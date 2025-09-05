package main

import (
	"context"
	"encoding/csv"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"time"

	"github.com/Wsine/feishu2md/core"
	"github.com/Wsine/feishu2md/utils"
	"github.com/chyroc/lark"
)

type fieldInfo struct {
	id   string
	name string
	typ  int64
	prop *lark.GetBitableFieldListRespItemProperty
}

// exportBitable exports a Feishu Bitable view to CSV or XLSX
// url must contain table=tbl... and may contain view=vew...
func exportBitable(ctx context.Context, client *core.Client, url string, format string, outputDir string, preferName string) error {
	// Extract tbl/vew from URL
	tableID, viewID := utils.ExtractBitableParams(url)
	if tableID == "" {
		return fmt.Errorf("bitable export requires query param 'table=tbl...' in URL")
	}

	// Resolve app token from wiki/docx page
	appToken, err := resolveBitableAppToken(ctx, client, url, tableID)
	if err != nil {
		return err
	}

	// Fetch names for file naming
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

	// Prepare field list in view order
	var viewPtr *string
	if viewID != "" {
		viewPtr = &viewID
	}
	fields, err := client.GetBitableFieldList(ctx, appToken, tableID, viewPtr)
	if err != nil {
		return fmt.Errorf("get fields failed: %w", err)
	}
	if len(fields) == 0 {
		return fmt.Errorf("no fields returned for table %s", tableID)
	}

	// Build mapping for select options by field id
	// fieldInfo is declared at package level for reuse in helpers
	ordered := make([]fieldInfo, 0, len(fields))
	for _, f := range fields {
		ordered = append(ordered, fieldInfo{
			id:   f.FieldID,
			name: f.FieldName,
			typ:  f.Type,
			prop: f.Property,
		})
	}

	// Pull records (paged)
	pageSize := int64(500)
	var pageToken *string
	rows := make([][]string, 0, 1024)

	for {
		resp, err := client.GetBitableRecordPage(ctx, appToken, tableID, viewPtr, pageToken, pageSize)
		if err != nil {
			return fmt.Errorf("list records failed: %w", err)
		}
		for _, item := range resp.Items {
			row := make([]string, 0, len(ordered))
			isCSV := strings.EqualFold(format, "csv")
			for _, col := range ordered {
				val := extractField(item.Fields, col.id, col.name)
				row = append(row, formatFieldValue(col, val, isCSV))
			}
			rows = append(rows, row)
		}
		if !resp.HasMore || resp.PageToken == "" {
			break
		}
		pageToken = &resp.PageToken
	}

	// Compose header
	headers := make([]string, 0, len(ordered))
	for _, col := range ordered {
		headers = append(headers, col.name)
	}

	// Build file name: App_Table_View (align with web export)
	parts := []string{sanitizeFileName(appName), sanitizeFileName(tableName)}
	if viewName != "" {
		parts = append(parts, sanitizeFileName(viewName))
	}
	baseName := strings.Join(parts, "_")

	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return err
	}

	switch strings.ToLower(format) {
	case "csv":
		out := filepath.Join(outputDir, baseName+".csv")
		if err := writeCSV(out, headers, rows); err != nil {
			return err
		}
		fmt.Printf("Exported CSV to %s\n", out)
	case "xlsx":
		out := filepath.Join(outputDir, baseName+".xlsx")
		if err := writeXLSX(out, headers, rows, ordered); err != nil {
			return err
		}
		fmt.Printf("Exported XLSX to %s\n", out)
	default:
		return fmt.Errorf("unsupported export format: %s", format)
	}

	return nil
}

// resolveBitableAppToken attempts to obtain the bitable app token (bascn...) from a given URL.
// It supports:
//   - wiki page embedding a bitable block: parse docx blocks to find Bitable tokens and probe using table id
//   - wiki link targeting a bitable file directly: wiki node obj_type == bitable
func resolveBitableAppToken(ctx context.Context, client *core.Client, url string, tableID string) (string, error) {
	// Validate and maybe resolve wiki object
	docType, docToken, err := utils.ValidateDocumentURL(url)
	if err != nil {
		return "", err
	}
	if docType == "wiki" {
		node, err := client.GetWikiNodeInfo(ctx, docToken)
		if err != nil {
			return "", err
		}
		// Direct bitable file
		if node.ObjType == "bitable" {
			return node.ObjToken, nil
		}
		// For wiki pages, traverse docx to find bitable blocks
		docType = node.ObjType
		docToken = node.ObjToken
	}

	// If we have a docx document, parse blocks to find bitable tokens
	if docType == "docx" {
		_, blocks, err := client.GetDocxContent(ctx, docToken)
		if err != nil {
			return "", err
		}
		// Map id -> block for traversal
		blockMap := map[string]*lark.DocxBlock{}
		for _, b := range blocks {
			blockMap[b.BlockID] = b
		}
		// Depth-first traversal to find bitable blocks
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
		// Find root
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
		// Remove duplicates while preserving order
		seen := map[string]bool{}
		uniq := make([]string, 0, len(tokens))
		for _, t := range tokens {
			if !seen[t] {
				seen[t] = true
				uniq = append(uniq, t)
			}
		}
		// If we have candidate tokens, pick the one that can access the target table
		if len(uniq) > 0 {
			for _, t := range uniq {
				if _, err := client.GetBitableFieldList(ctx, t, tableID, nil); err == nil {
					return t, nil
				}
			}
			// fallback to first
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
	// Write UTF-8 BOM for better Excel compatibility on Windows
	f.Write([]byte{0xEF, 0xBB, 0xBF})
	w := csv.NewWriter(f)
	// Keep Excel-friendly defaults
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
	// Lazy import to avoid adding heavy dependency if not used in build paths
	// We rely on excelize being in go.mod.
	type excelWriter interface {
		SetCellValue(string, string, interface{}) error
		NewSheet(string) int
		SetActiveSheet(int)
		SaveAs(string) error
		DeleteSheet(string)
	}

	// Inline minimal use of excelize via reflection-unsafe import would be messy; instead, we
	// directly import excelize here.
	return writeXLSXWithExcelize(path, headers, rows, fields)
}

// Separated into its own func to keep the main flow simple
func writeXLSXWithExcelize(path string, headers []string, rows [][]string, fields []fieldInfo) error {
	// Import here
	// nolint:depguard
	f := excelizeNewFile()
	sheet := "Sheet1"
	idx := f.NewSheet(sheet)
	// header
	for i, h := range headers {
		cell := excelColumnName(i+1) + "1"
		_ = f.SetCellValue(sheet, cell, h)
	}
	// data
	for rIdx, r := range rows {
		for cIdx, v := range r {
			cell := excelColumnName(cIdx+1) + fmt.Sprintf("%d", rIdx+2)
			_ = f.SetCellValue(sheet, cell, v)
		}
	}

	// Add data validation for select fields
	for colIdx, field := range fields {
		if shouldAddDropdown(field) {
			options := getFieldOptions(field)
			if len(options) > 0 {
				// Create range for the entire column (excluding header)
				colName := excelColumnName(colIdx + 1)
				rangeAddr := fmt.Sprintf("%s2:%s1048576", colName, colName) // Excel max rows

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
	// Delete default if different
	if sheet != "Sheet1" {
		f.DeleteSheet("Sheet1")
	}
	return f.SaveAs(path)
}

// Helpers for excelize (wrapped to avoid direct import names in unit).
//go:generate echo "placeholder"

// excelize minimal shims
type excelFile interface {
	NewSheet(name string) int
	SetCellValue(sheet, axis string, value interface{}) error
	SetActiveSheet(index int)
	DeleteSheet(name string)
	SaveAs(name string) error
	AddDataValidation(sheet string, dv DataValidation) error
}

// DataValidation represents Excel data validation
type DataValidation interface {
	SetRange(rangeAddr string) error
	SetDropList([]string) error
}

func excelizeNewFile() excelFile {
	// Import inside to keep the rest of the code independent
	return excelizeNew()
}

// shouldAddDropdown determines if a field should have a dropdown
func shouldAddDropdown(field fieldInfo) bool {
	// Type 3: single select, Type 4: multi select
	return field.typ == 3 || field.typ == 4
}

// getFieldOptions extracts dropdown options from field properties
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

// createDataValidation creates a new data validation instance
func createDataValidation() DataValidation {
	return newDataValidation()
}

// Implemented in xlsx_shim.go using excelize

// extractField tries both field name and field id keys.
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
	// As a fallback, try case-insensitive name match (some locales)
	lower := strings.ToLower(fieldName)
	for k, v := range m {
		if strings.ToLower(k) == lower {
			return v
		}
	}
	return nil
}

func formatFieldValue(col fieldInfo, v interface{}, isCSV bool) string {
	if v == nil {
		return ""
	}
	// Normalize by type when possible
	switch col.typ {
	case 1: // text
		switch t := v.(type) {
		case string:
			return t
		case []interface{}:
			parts := make([]string, 0, len(t))
			for _, it := range t {
				parts = append(parts, fmt.Sprint(it))
			}
			return strings.Join(parts, "\n")
		default:
			return fmt.Sprint(v)
		}
	case 2: // number
		return fmt.Sprint(v)
	case 3: // single select
		return mapSelectOptionName(col.prop, v)
	case 4: // multi select
		return joinList(v, func(x interface{}) string { return mapSelectOptionName(col.prop, x) })
	case 5: // date/time
		return fmt.Sprint(v)
	case 7: // checkbox
		return fmt.Sprint(v)
	case 11: // user
		return joinList(v, func(x interface{}) string { return pickStringField(x, "name") })
	case 15: // url
		if s := pickStringField(v, "link"); s != "" {
			return s
		}
		return fmt.Sprint(v)
	case 17: // attachment
		return joinList(v, func(x interface{}) string { return pickStringField(x, "name") })
	case 18: // relation (single)
		if isCSV {
			return ""
		} // 对齐飞书 CSV 导出行为
		// Handle both single map and array of maps
		return joinList(v, func(x interface{}) string {
			if m, ok := x.(map[string]interface{}); ok {
				// Debug: print the structure to understand the data
				// fmt.Printf("DEBUG relation field: %+v\n", m)

				// Check if text field exists and use it directly
				if s, ok := m["text"].(string); ok && s != "" {
					return s
				}

				// Then check text_arr
				if arr, ok := m["text_arr"].([]interface{}); ok && len(arr) > 0 {
					parts := make([]string, 0, len(arr))
					for _, t := range arr {
						parts = append(parts, fmt.Sprint(t))
					}
					return strings.Join(parts, ",")
				}
				if ids, ok := m["record_ids"].([]interface{}); ok && len(ids) > 0 {
					// fallback to first record id
					return fmt.Sprint(ids[0])
				}

				// 检查是否有特殊的type字段指示这是一个文本类型的关联
				if typ, ok := m["type"].(string); ok && typ == "text" {
					// 如果是text类型但没有实际文本内容，可能需要通过其他方式获取
					// 先返回空，但这可能不是最终解决方案
					return ""
				}

				// 对于空的关联字段，返回空字符串而不是整个map结构
				return ""
			}
			return fmt.Sprint(x)
		})
	case 21: // bi-direction relation
		if isCSV {
			return ""
		}
		return joinList(v, func(x interface{}) string {
			if m, ok := x.(map[string]interface{}); ok {
				if arr, ok := m["text_arr"].([]interface{}); ok && len(arr) > 0 {
					parts := make([]string, 0, len(arr))
					for _, t := range arr {
						parts = append(parts, fmt.Sprint(t))
					}
					return strings.Join(parts, ",")
				}
				if s, ok := m["text"].(string); ok && s != "" {
					return s
				}
				// 对于空的关联字段，返回空字符串而不是整个map结构
				return ""
			}
			return fmt.Sprint(x)
		})
	default:
		// unknown or complex types: best effort string
		return fmt.Sprint(v)
	}
}

func mapSelectOptionName(prop *lark.GetBitableFieldListRespItemProperty, v interface{}) string {
	id := fmt.Sprint(v)
	// If v is map with id field
	if m, ok := v.(map[string]interface{}); ok {
		if s, ok := m["id"].(string); ok {
			id = s
		} else if s, ok := m["name"].(string); ok {
			return s
		}
	}
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

func pickStringField(v interface{}, key string) string {
	if m, ok := v.(map[string]interface{}); ok {
		if s, ok := m[key].(string); ok {
			return s
		}
		// Some user objects may have nested fields
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
		// Handle single item case
		if v == nil {
			return ""
		}
		// try map of id->value sorted by key
		if m, ok := v.(map[string]interface{}); ok {
			keys := make([]string, 0, len(m))
			for k := range m {
				keys = append(keys, k)
			}
			sort.Strings(keys)
			out := make([]string, 0, len(keys))
			for _, k := range keys {
				mapped := mapFn(m[k])
				if mapped != "" {
					out = append(out, mapped)
				}
			}
			return strings.Join(out, ",")
		}
		return mapFn(v)
	}
}

func sanitizeFileName(name string) string {
	// Replace invalid characters for Windows and general FS
	invalid := regexp.MustCompile(`[\\/:*?"<>|]`)
	s := invalid.ReplaceAllString(name, "_")
	s = strings.TrimSpace(s)
	if s == "" {
		s = fmt.Sprintf("export_%d", time.Now().Unix())
	}
	return s
}

func excelColumnName(n int) string {
	// 1 -> A, 27 -> AA
	name := ""
	for n > 0 {
		n--
		name = string('A'+(n%26)) + name
		n /= 26
	}
	return name
}
