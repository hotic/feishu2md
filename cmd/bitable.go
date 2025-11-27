package main

import (
	"context"
	"encoding/csv"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/Wsine/feishu2md/core"
	"github.com/Wsine/feishu2md/utils"
	"github.com/chyroc/lark"
)

// 字段元信息(按视图顺序排列)
type fieldInfo struct {
	id   string
	name string
	typ  int64
	prop *lark.GetBitableFieldListRespItemProperty
}

// 导出多维表格为 CSV/XLSX
// url 必须包含 table=tbl...;若包含 view=vew... 将按视图顺序组织列
// preferName 为空时,文件名采用 App_表_视图;否则使用自定义名称
// viewFieldsOnly 为 true 时,仅导出该视图中"可见"的字段(尽量贴近 Web 导出)
// filterImages 为 true 时,过滤掉图片文件引用,减少无用文本噪音
// 返回生成文件的实际文件名
func exportBitable(ctx context.Context, client *core.Client, url string, format string, outputDir string, preferName string, viewFieldsOnly bool, filterImages bool) (string, error) {
	// 从 URL 提取 tbl/vew 参数
	tableID, viewID := utils.ExtractBitableParams(url)
	if tableID == "" {
		return "", fmt.Errorf("bitable export requires query param 'table=tbl...' in URL")
	}

	// 从 wiki/docx 页面解析 app token
	appToken, err := resolveBitableAppToken(ctx, client, url, tableID)
	if err != nil {
		return "", err
	}

	// 获取应用、表、视图名称用于文件命名
	appName := "bitable"
	if app, err := client.GetBitableMeta(ctx, appToken); err == nil && app != nil {
		if app.Name != "" {
			appName = app.Name
		}
	}
	tableName := tableID
	if tables, err := client.GetBitableTableList(ctx, appToken); err == nil {
		for _, t := range tables {
			if t.TableID == tableID {
				if t.Name != "" {
					tableName = t.Name
				}
				break
			}
		}
	}
	viewName := viewID
	if viewID != "" {
		if views, err := client.GetBitableViewList(ctx, appToken, tableID); err == nil {
			for _, v := range views {
				if v.ViewID == viewID {
					if v.ViewName != "" {
						viewName = v.ViewName
					}
					break
				}
			}
		}
	}

	// 按视图顺序准备字段列表
	var viewPtr *string
	if viewID != "" {
		viewPtr = &viewID
	}
	fields, err := client.GetBitableFieldList(ctx, appToken, tableID, viewPtr)
	if err != nil {
		return "", fmt.Errorf("get fields failed: %w", err)
	}
	if len(fields) == 0 {
		return "", fmt.Errorf("no fields returned for table %s", tableID)
	}

	// 构建字段信息映射(用于选项字段的名称映射)
	ordered := make([]fieldInfo, 0, len(fields))
	for _, f := range fields {
		ordered = append(ordered, fieldInfo{
			id:   f.FieldID,
			name: f.FieldName,
			typ:  f.Type,
			prop: f.Property,
		})
	}

	// 默认隐藏系统字段以模拟飞书Web导出行为
	if !isTruthy(os.Getenv("FEISHU2MD_INCLUDE_SYSTEM_FIELDS")) {
		filtered := make([]fieldInfo, 0, len(ordered))
		for _, c := range ordered {
			if c.typ == 1001 || c.typ == 1002 || c.typ == 1003 || c.typ == 1004 || c.typ == 1005 {
				continue
			}
			filtered = append(filtered, c)
		}
		ordered = filtered
	}

	// 分页拉取记录
	pageSize := int64(500)
	var pageToken *string
	rows := make([][]string, 0, 1024)
	appliedVisible := false

	for {
		resp, err := client.GetBitableRecordPage(ctx, appToken, tableID, viewPtr, pageToken, pageSize)
		if err != nil {
			return "", fmt.Errorf("list records failed: %w", err)
		}

		// 根据视图实际可见字段缩小列范围(基于记录中的实际字段键)
		if viewFieldsOnly && !appliedVisible {
			visible := map[string]bool{}
			for _, it := range resp.Items {
				for k := range it.Fields {
					visible[strings.ToLower(k)] = true
				}
			}
			if len(visible) > 0 {
				filtered := make([]fieldInfo, 0, len(ordered))
				for _, c := range ordered {
					if visible[strings.ToLower(c.name)] || visible[strings.ToLower(c.id)] {
						filtered = append(filtered, c)
					}
				}
				if len(filtered) > 0 {
					ordered = filtered
				}
			}
			appliedVisible = true
		}

		for _, item := range resp.Items {
			row := make([]string, 0, len(ordered))
			isCSV := strings.EqualFold(format, "csv")
			for _, col := range ordered {
				val := extractField(item.Fields, col.id, col.name)
				row = append(row, formatFieldValue(col, val, isCSV, filterImages))
			}
			rows = append(rows, row)
		}
		if !resp.HasMore || resp.PageToken == "" {
			break
		}
		pageToken = &resp.PageToken
	}

	// 组装表头
	headers := make([]string, 0, len(ordered))
	for _, col := range ordered {
		headers = append(headers, col.name)
	}

	// 构建文件名:根据preferName决定使用自定义名称还是原始标题
	var baseName string
	if preferName != "" {
		// 使用配置中的自定义名称
		baseName = sanitizeFileName(preferName)
	} else {
		// 使用原始标题: App_Table_View (对齐Web导出)
		parts := []string{sanitizeFileName(appName), sanitizeFileName(tableName)}
		if viewName != "" {
			parts = append(parts, sanitizeFileName(viewName))
		}
		baseName = strings.Join(parts, "_")
	}

	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return "", err
	}

	var actualFileName string
	switch strings.ToLower(format) {
	case "csv":
		actualFileName = baseName + ".csv"
		out := filepath.Join(outputDir, actualFileName)
		if err := writeCSV(out, headers, rows); err != nil {
			return "", err
		}
		fmt.Printf("Exported CSV to %s\n", out)
	case "xlsx":
		actualFileName = baseName + ".xlsx"
		out := filepath.Join(outputDir, actualFileName)
		if err := writeXLSX(out, headers, rows, ordered); err != nil {
			return "", err
		}
		fmt.Printf("Exported XLSX to %s\n", out)
	default:
		return "", fmt.Errorf("unsupported export format: %s", format)
	}

	return actualFileName, nil
}

// resolveBitableAppToken 尝试从给定 URL 获取多维表格的 app token (bascn...)
// 支持:
//   - 嵌入多维表格块的 wiki 页面:解析 docx 块以查找 Bitable token 并使用 table id 探测
//   - 直接指向多维表格文件的 wiki 链接:wiki 节点 obj_type == bitable
func resolveBitableAppToken(ctx context.Context, client *core.Client, url string, tableID string) (string, error) {
	// 验证并可能解析 wiki 对象
	docType, docToken, err := utils.ValidateDocumentURL(url)
	if err != nil {
		return "", err
	}
	if docType == "wiki" {
		node, err := client.GetWikiNodeInfo(ctx, docToken)
		if err != nil {
			return "", err
		}
		// 直接的多维表格文件
		if node.ObjType == "bitable" {
			return node.ObjToken, nil
		}
		// 对于 wiki 页面,遍历 docx 以查找多维表格块
		docType = node.ObjType
		docToken = node.ObjToken
	}

	// 如果我们有 docx 文档,解析块以查找多维表格 token
	if docType == "docx" {
		_, blocks, err := client.GetDocxContent(ctx, docToken)
		if err != nil {
			return "", err
		}
		// 构建 id -> block 映射以便遍历
		blockMap := map[string]*lark.DocxBlock{}
		for _, b := range blocks {
			blockMap[b.BlockID] = b
		}
		// 深度优先遍历查找多维表格块
		var tokens []string
		var visit func(id string)
		visit = func(id string) {
			b := blockMap[id]
			if b == nil {
				return
			}
			if b.BlockType == lark.DocxBlockTypeBitable && b.Bitable != nil && b.Bitable.Token != "" {
				tokens = append(tokens, b.Bitable.Token)
			}
			for _, c := range b.Children {
				visit(c)
			}
		}
		// 查找根节点
		var root *lark.DocxBlock
		for _, b := range blocks {
			if b.ParentID == "" || b.ParentID == b.BlockID {
				root = b
				break
			}
		}
		if root != nil {
			visit(root.BlockID)
		} else {
			for _, b := range blocks {
				visit(b.BlockID)
			}
		}
		// 去重并保持顺序
		seen := map[string]bool{}
		uniq := make([]string, 0, len(tokens))
		for _, t := range tokens {
			if !seen[t] {
				seen[t] = true
				uniq = append(uniq, t)
			}
		}
		// 如果我们有候选 token,选择能访问目标表的那个
		if len(uniq) > 0 {
			for _, t := range uniq {
				if _, err := client.GetBitableFieldList(ctx, t, tableID, nil); err == nil {
					return t, nil
				}
			}
			// 后备方案:使用第一个
			return uniq[0], nil
		}
	}

	return "", errors.New("failed to resolve bitable app token from URL; ensure the page contains an embedded table or point to a bitable file")
}

func writeCSV(path string, headers []string, rows [][]string) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	// 写入 UTF-8 BOM 以提高 Windows Excel 兼容性
	f.Write([]byte{0xEF, 0xBB, 0xBF})
	w := csv.NewWriter(f)
	// 保持 Excel 友好的默认设置
	if err := w.Write(headers); err != nil {
		return err
	}
	for _, r := range rows {
		if err := w.Write(r); err != nil {
			return err
		}
	}
	w.Flush()
	return w.Error()
}

func writeXLSX(path string, headers []string, rows [][]string, fields []fieldInfo) error {
	// 延迟导入以避免在不使用时增加重量级依赖
	// 我们依赖 go.mod 中的 excelize
	return writeXLSXWithExcelize(path, headers, rows, fields)
}

// 分离到单独的函数以保持主流程简洁
func writeXLSXWithExcelize(path string, headers []string, rows [][]string, fields []fieldInfo) error {
	f := excelizeNewFile()
	sheet := "Sheet1"
	idx := f.NewSheet(sheet)
	// 写入表头
	for i, h := range headers {
		cell := excelColumnName(i+1) + "1"
		_ = f.SetCellValue(sheet, cell, h)
	}
	// 写入数据
	for rIdx, r := range rows {
		for cIdx, v := range r {
			cell := excelColumnName(cIdx+1) + fmt.Sprintf("%d", rIdx+2)
			_ = f.SetCellValue(sheet, cell, v)
		}
	}

	// 为选择字段添加数据验证
	for colIdx, field := range fields {
		if shouldAddDropdown(field) {
			options := getFieldOptions(field)
			if len(options) > 0 {
				// 为整个列创建范围(排除表头)
				colName := excelColumnName(colIdx + 1)
				rangeAddr := fmt.Sprintf("%s2:%s1048576", colName, colName) // Excel 最大行数

				dv := createDataValidation()
				if err := dv.SetRange(rangeAddr); err == nil {
					if err := dv.SetDropList(options); err == nil {
						_ = f.AddDataValidation(sheet, dv)
					}
				}
			}
		}
	}

	f.SetActiveSheet(idx)
	// 如果不同则删除默认工作表
	if sheet != "Sheet1" {
		f.DeleteSheet("Sheet1")
	}
	return f.SaveAs(path)
}

// excelize 最小封装
// 在 xlsx_shim.go 中使用 excelize 实现

// excelize 文件接口
type excelFile interface {
	NewSheet(name string) int
	SetCellValue(sheet, axis string, value interface{}) error
	SetActiveSheet(index int)
	DeleteSheet(name string)
	SaveAs(name string) error
	AddDataValidation(sheet string, dv DataValidation) error
}

// DataValidation 表示 Excel 数据验证
type DataValidation interface {
	SetRange(rangeAddr string) error
	SetDropList([]string) error
}

func excelizeNewFile() excelFile {
	// 在内部导入以保持其余代码独立
	return excelizeNew()
}

// 判断字段是否应该有下拉菜单
func shouldAddDropdown(field fieldInfo) bool {
	// Type 3: 单选, Type 4: 多选
	return field.typ == 3 || field.typ == 4
}

// 从字段属性中提取下拉选项
func getFieldOptions(field fieldInfo) []string {
	if field.prop == nil || len(field.prop.Options) == 0 {
		return nil
	}

	options := make([]string, 0, len(field.prop.Options))
	for _, opt := range field.prop.Options {
		if opt.Name != "" {
			options = append(options, opt.Name)
		}
	}
	return options
}

// 创建新的数据验证实例
func createDataValidation() DataValidation {
	return newDataValidation()
}

// isImageFile 检测文件名是否为图片文件
func isImageFile(filename string) bool {
	if filename == "" {
		return false
	}
	lower := strings.ToLower(filename)
	imageExts := []string{".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".svg", ".ico", ".tiff", ".tif"}
	for _, ext := range imageExts {
		if strings.HasSuffix(lower, ext) {
			return true
		}
	}
	return false
}

// 尝试使用字段名和字段 ID 键提取字段值
func extractField(m map[string]interface{}, fieldID, fieldName string) interface{} {
	if m == nil {
		return nil
	}
	if v, ok := m[fieldName]; ok {
		return v
	}
	if v, ok := m[fieldID]; ok {
		return v
	}
	// 后备方案:尝试不区分大小写的名称匹配(某些语言环境)
	lower := strings.ToLower(fieldName)
	for k, v := range m {
		if strings.ToLower(k) == lower {
			return v
		}
	}
	return nil
}

// 格式化字段值为字符串表示
func formatFieldValue(col fieldInfo, v interface{}, isCSV bool, filterImages bool) string {
	if v == nil {
		return ""
	}

	// 预处理：如果是 map[type:xxx value:xxx] 格式，先提取value
	// 这是 textFieldAsArray=true 和 displayFormulaRef=true 时的常见格式
	if m, ok := v.(map[string]interface{}); ok {
		if typeStr, hasType := m["type"].(string); hasType {
			if value, hasValue := m["value"]; hasValue {
				// 对于某些引用字段，如果有value_extra则显示为空
				if (typeStr == "single_option" || typeStr == "multi_option") && m["value_extra"] != nil {
					return ""
				}
				// 提取value继续处理
				v = value
				// 如果value是单元素数组且包含attachmentToken，说明是附件字段，直接提取第一个元素
				if arr, isArr := v.([]interface{}); isArr && len(arr) == 1 {
					if itemMap, ok := arr[0].(map[string]interface{}); ok {
						if _, hasToken := itemMap["attachmentToken"]; hasToken {
							// 直接返回文件名
							if name, ok := itemMap["name"].(string); ok {
								// 如果开启了图片过滤且是图片文件，返回空
								if filterImages && isImageFile(name) {
									return ""
								}
								return name
							}
						}
					}
				}
			}
		}
	}

	// 提前检查:如果值是包含 type:single_option 或 type:multi_option 的 map
	// 并且还包含 value_extra,这通常是引用字段的原始数据结构
	// 在官方导出中,这些字段显示为空(因为它们是反向引用或查找字段)
	if m, ok := v.(map[string]interface{}); ok {
		if typ, ok := m["type"].(string); ok {
			if (typ == "single_option" || typ == "multi_option") && m["value_extra"] != nil {
				// 这是引用字段的完整数据结构,应该显示为空
				// (与官方导出行为一致)
				return ""
			}
		}
	}

	// 根据类型进行规范化处理
	switch col.typ {
	case 1: // 文本
		// 当 text_field_as_array=true 时,文本字段可能返回 map[text:xxx type:text]
		if m, ok := v.(map[string]interface{}); ok {
			if s, ok := m["text"].(string); ok && s != "" {
				// 如果开启了图片过滤且是图片文件，返回空
				if filterImages && isImageFile(s) {
					return ""
				}
				return s
			}
			// 尝试提取name字段作为后备
			if s, ok := m["name"].(string); ok && s != "" {
				// 如果开启了图片过滤且是图片文件，返回空
				if filterImages && isImageFile(s) {
					return ""
				}
				return s
			}
			// 如果是空map或无法提取有用信息，返回空字符串
			return ""
		}
		switch t := v.(type) {
		case string:
			// 如果开启了图片过滤且是图片文件，返回空
			if filterImages && isImageFile(t) {
				return ""
			}
			return t
		case []interface{}:
			parts := make([]string, 0, len(t))
			for _, it := range t {
				// 每个项也可能是带有 text 字段的 map
				if m, ok := it.(map[string]interface{}); ok {
					if s, ok := m["text"].(string); ok {
						// 如果开启了图片过滤且是图片文件，跳过
						if filterImages && isImageFile(s) {
							continue
						}
						parts = append(parts, s)
						continue
					}
					if s, ok := m["name"].(string); ok {
						// 如果开启了图片过滤且是图片文件，跳过
						if filterImages && isImageFile(s) {
							continue
						}
						parts = append(parts, s)
						continue
					}
					// 跳过无法提取有用信息的map
					continue
				}
				s := fmt.Sprint(it)
				// 如果开启了图片过滤且是图片文件，跳过
				if filterImages && isImageFile(s) {
					continue
				}
				if s != "" && s != "map[]" {
					parts = append(parts, s)
				}
			}
			return strings.Join(parts, "\n")
		default:
			// 对于未知类型，返回空字符串而不是fmt.Sprint
			return ""
		}
	case 2: // 数字
		return fmt.Sprint(v)
	case 3: // 单选
		return mapSelectOptionName(col.prop, v)
	case 4: // 多选
		return joinList(v, func(x interface{}) string { return mapSelectOptionName(col.prop, x) })
	case 5: // 日期/时间
		if s := formatTimeValue(v); s != "" {
			return s
		}
		return fmt.Sprint(v)
	case 7: // 复选框
		return fmt.Sprint(v)
	case 11: // 人员
		return joinList(v, func(x interface{}) string { return pickStringField(x, "name") })
	case 15: // 链接
		if s := pickStringField(v, "link"); s != "" {
			return s
		}
		return fmt.Sprint(v)
	case 17: // 附件
		return joinList(v, func(x interface{}) string {
			name := pickStringField(x, "name")
			// 如果开启了图片过滤且是图片文件，返回空字符串（会被 joinList 过滤掉）
			if filterImages && isImageFile(name) {
				return ""
			}
			return name
		})
	case 18: // 单向关联(从当前表指向其他表)
		// 官方导出中,单向关联字段显示为空(因为它是反向引用)
		// 如果字段有实际的文本内容,则显示;否则返回空
		if m, ok := v.(map[string]interface{}); ok {
			// 检查是否是空的关联字段(type=text 但没有实际内容)
			if typ, ok := m["type"].(string); ok && typ == "text" {
				// 如果有 text 字段且非空,则显示
				if text, ok := m["text"].(string); ok && text != "" {
					return text
				}
				// 如果有 text_arr 且非空,则显示
				if textArr, ok := m["text_arr"].([]interface{}); ok && len(textArr) > 0 {
					parts := make([]string, 0, len(textArr))
					for _, t := range textArr {
						if s := fmt.Sprint(t); s != "" {
							parts = append(parts, s)
						}
					}
					if len(parts) > 0 {
						return strings.Join(parts, ",")
					}
				}
				// 空的反向引用字段,返回空字符串
				return ""
			}
		}
		// 处理数组形式的关联
		return joinList(v, func(x interface{}) string {
			if m, ok := x.(map[string]interface{}); ok {
				if s, ok := m["text"].(string); ok && s != "" {
					return s
				}
				if arr, ok := m["text_arr"].([]interface{}); ok && len(arr) > 0 {
					parts := make([]string, 0, len(arr))
					for _, t := range arr {
						parts = append(parts, fmt.Sprint(t))
					}
					return strings.Join(parts, ",")
				}
				// 空的关联,返回空字符串
				return ""
			}
			return fmt.Sprint(x)
		})
	case 19: // 查找引用字段
		// Type 19 字段在 display_formula_ref=true 时返回结构化数据
		if m, ok := v.(map[string]interface{}); ok {
			// 检查 value_extra.options (来自 display_formula_ref=true)
			if valueExtra, ok := m["value_extra"].(map[string]interface{}); ok {
				if options, ok := valueExtra["options"].([]interface{}); ok && len(options) > 0 {
					// 从所有选项中提取名称
					names := make([]string, 0, len(options))
					for _, opt := range options {
						if optMap, ok := opt.(map[string]interface{}); ok {
							if name, ok := optMap["name"].(string); ok && name != "" {
								names = append(names, name)
							}
						}
					}
					if len(names) > 0 {
						return strings.Join(names, ",")
					}
				}
			}

			// 检查是否为单选类型(single_option)并返回空字符串
			// 这与官方导出行为一致:引用字段在官方导出中显示为空
			if typ, ok := m["type"].(string); ok {
				if typ == "single_option" || typ == "multi_option" {
					// 这是引用其他表的字段,在官方导出中显示为空
					return ""
				}
			}

			// 后备方案:检查 text 字段(在 text_field_as_array=true 时常见)
			if s, ok := m["text"].(string); ok && s != "" {
				return s
			}
			// 检查 name 字段
			if s, ok := m["name"].(string); ok && s != "" {
				return s
			}
		}
		// 如果是简单字符串(如选项 ID),尝试映射它
		return mapSelectOptionName(col.prop, v)
	case 21: // 双向关联(两个表之间的双向引用)
		// 官方导出中,双向关联字段通常显示为空(因为它们是反向引用)
		// 只有在有实际文本内容时才显示
		return joinList(v, func(x interface{}) string {
			if m, ok := x.(map[string]interface{}); ok {
				if arr, ok := m["text_arr"].([]interface{}); ok && len(arr) > 0 {
					parts := make([]string, 0, len(arr))
					for _, t := range arr {
						if s := fmt.Sprint(t); s != "" {
							parts = append(parts, s)
						}
					}
					if len(parts) > 0 {
						return strings.Join(parts, ",")
					}
				}
				if s, ok := m["text"].(string); ok && s != "" {
					return s
				}
				// 空的双向关联,返回空字符串
				return ""
			}
			return ""
		})
	case 1001, 1002: // 创建/更新时间
		if s := formatTimeValue(v); s != "" {
			return s
		}
		return fmt.Sprint(v)
	case 1003, 1004: // 创建者/修改者(人员)
		return joinList(v, func(x interface{}) string { return pickStringField(x, "name") })
	default:
		// 未知或复杂类型:尝试从map中提取有用信息,否则转换为字符串
		// 首先尝试从map[text:xxx type:text]结构中提取text字段
		if m, ok := v.(map[string]interface{}); ok {
			// 优先提取text字段
			if s, ok := m["text"].(string); ok && s != "" {
				return s
			}
			// 其次尝试name字段（对于附件等）
			if s, ok := m["name"].(string); ok && s != "" {
				return s
			}
			// 对于数组,递归处理
			if arr, ok := m["text_arr"].([]interface{}); ok && len(arr) > 0 {
				parts := make([]string, 0, len(arr))
				for _, item := range arr {
					if s := fmt.Sprint(item); s != "" {
						parts = append(parts, s)
					}
				}
				if len(parts) > 0 {
					return strings.Join(parts, ",")
				}
			}
		}
		// 处理数组类型:尝试提取每个元素的有用信息
		if arr, ok := v.([]interface{}); ok && len(arr) > 0 {
			parts := make([]string, 0, len(arr))
			for _, item := range arr {
				// 对每个元素递归处理
				if m, ok := item.(map[string]interface{}); ok {
					// 如果有text字段，优先使用
					if s, ok := m["text"].(string); ok && s != "" {
						parts = append(parts, s)
						continue
					}
					// 如果有name字段（附件、人员等），使用name
					if s, ok := m["name"].(string); ok && s != "" {
						parts = append(parts, s)
						continue
					}
					// 如果有attachmentToken字段，说明是附件，只提取name
					if _, hasToken := m["attachmentToken"]; hasToken {
						if s, ok := m["name"].(string); ok && s != "" {
							// 如果开启了图片过滤且是图片文件，跳过
							if filterImages && isImageFile(s) {
								continue
							}
							parts = append(parts, s)
							continue
						}
						// 跳过无法提取name的附件
						continue
					}
				}
				// 对于不是map的元素，转换为字符串
				if s := fmt.Sprint(item); s != "" && s != "map[]" {
					parts = append(parts, s)
				}
			}
			if len(parts) > 0 {
				return strings.Join(parts, ",")
			}
		}
		// 最后才使用fmt.Sprint
		result := fmt.Sprint(v)
		// 如果结果是空的map,返回空字符串
		if result == "map[]" || result == "[]" {
			return ""
		}
		return result
	}
}

// 将选项ID映射到选项名称
func mapSelectOptionName(prop *lark.GetBitableFieldListRespItemProperty, v interface{}) string {
	id := fmt.Sprint(v)
	// 如果 v 是带有 id 字段的 map
	if m, ok := v.(map[string]interface{}); ok {
		if s, ok := m["id"].(string); ok {
			id = s
		} else if s, ok := m["name"].(string); ok {
			return s
		}
	}

	// 移除括号(例如, "[optXXXXX]" -> "optXXXXX")
	id = strings.TrimPrefix(id, "[")
	id = strings.TrimSuffix(id, "]")

	if prop != nil && len(prop.Options) > 0 {
		for _, opt := range prop.Options {
			if opt.ID == id {
				if opt.Name != "" {
					return opt.Name
				}
				break
			}
		}
	}
	return id
}

// 从对象中提取指定键的字符串字段
func pickStringField(v interface{}, key string) string {
	if m, ok := v.(map[string]interface{}); ok {
		if s, ok := m[key].(string); ok {
			return s
		}
		// 某些用户对象可能有嵌套字段
		for k, vv := range m {
			if strings.EqualFold(k, key) {
				if s, ok := vv.(string); ok {
					return s
				}
			}
		}
	}
	return fmt.Sprint(v)
}

// 连接列表,对每个元素应用映射函数
func joinList(v interface{}, mapFn func(interface{}) string) string {
	switch arr := v.(type) {
	case []interface{}:
		if len(arr) == 0 {
			return ""
		}
		out := make([]string, 0, len(arr))
		for _, it := range arr {
			mapped := mapFn(it)
			if mapped != "" {
				out = append(out, mapped)
			}
		}
		return strings.Join(out, ",")
	default:
		// 处理单项情况
		if v == nil {
			return ""
		}
		// 对于单个map对象，直接应用mapFn（而不是遍历键值对）
		// 这对于附件等字段很重要，避免输出整个map结构
		return mapFn(v)
	}
}

// 清理文件名中的非法字符
func sanitizeFileName(name string) string {
	// 替换 Windows 和一般文件系统的非法字符
	invalid := regexp.MustCompile(`[\\/:*?"<>|]`)
	s := invalid.ReplaceAllString(name, "_")
	s = strings.TrimSpace(s)
	if s == "" {
		s = fmt.Sprintf("export_%d", time.Now().Unix())
	}
	return s
}

// 生成 Excel 列名(1 -> A, 27 -> AA)
func excelColumnName(n int) string {
	name := ""
	for n > 0 {
		n--
		name = string('A'+(n%26)) + name
		n /= 26
	}
	return name
}

// 检查常见的真值字符串("1","true","yes")
func isTruthy(s string) bool {
	switch strings.ToLower(strings.TrimSpace(s)) {
	case "1", "true", "yes", "y", "on":
		return true
	default:
		return false
	}
}

// 尝试将时间戳数字渲染为可读时间
// 支持秒或毫秒时间戳,也支持科学计数法字符串
func formatTimeValue(v interface{}) string {
	toInt := func(x interface{}) (int64, bool) {
		switch t := x.(type) {
		case int64:
			return t, true
		case int:
			return int64(t), true
		case float64:
			// 大的时间戳可能以浮点数形式出现;合理四舍五入
			return int64(t + 0.5), true
		case string:
			s := strings.TrimSpace(t)
			if s == "" {
				return 0, false
			}
			// 先尝试整数
			if n, err := strconv.ParseInt(s, 10, 64); err == nil {
				return n, true
			}
			// 后备方案:浮点数(处理科学计数法)
			if f, err := strconv.ParseFloat(s, 64); err == nil {
				return int64(f + 0.5), true
			}
			return 0, false
		default:
			return 0, false
		}
	}

	n, ok := toInt(v)
	if !ok || n == 0 {
		return ""
	}
	// 检测毫秒 vs 秒
	// 阈值:任何 > 10^11 的很可能是毫秒
	if n > 100_000_000_000 { // ~ 1973-03-03 毫秒阈值
		sec := n / 1000
		return time.Unix(sec, 0).Format("2006-01-02 15:04:05")
	}
	return time.Unix(n, 0).Format("2006-01-02 15:04:05")
}
