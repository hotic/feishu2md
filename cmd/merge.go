package main

import (
	"fmt"
	"html"
	"io/fs"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"time"

	"github.com/urfave/cli/v2"
)

type MergeOpts struct {
	inputDir   string
	outputDir  string
	filename   string
	configPath string
	original   bool
}

var mergeOpts = MergeOpts{}

// getMergeCommand returns the merge command definition
func getMergeCommand() *cli.Command {
	return &cli.Command{
		Name:  "merge",
		Usage: "Merge all .md files from input directory into a single file",
		Flags: []cli.Flag{
			&cli.StringFlag{
				Name:        "config",
				Aliases:     []string{"c"},
				Value:       "",
				Usage:       "Path to config file (defaults to config.yml in current directory)",
				Destination: &mergeOpts.configPath,
			},
			&cli.StringFlag{
				Name:        "input",
				Aliases:     []string{"i"},
				Value:       "",
				Usage:       "Input directory containing .md files to merge (overrides config)",
				Destination: &mergeOpts.inputDir,
			},
			&cli.StringFlag{
				Name:        "output",
				Aliases:     []string{"o"},
				Value:       "",
				Usage:       "Output directory for the merged file (overrides config)",
				Destination: &mergeOpts.outputDir,
			},
			&cli.StringFlag{
				Name:        "filename",
				Aliases:     []string{"f"},
				Value:       "",
				Usage:       "Name of the merged output file (overrides config)",
				Destination: &mergeOpts.filename,
			},
			&cli.BoolFlag{
				Name:        "original",
				Usage:       "Output original (uncompacted) content without token-compact processing",
				Value:       false,
				Destination: &mergeOpts.original,
			},
		},
		Action: func(ctx *cli.Context) error {
			return handleMergeCommand()
		},
	}
}

// handleMergeCommand processes the merge command
func handleMergeCommand() error {
	// 加载配置文件
	config, err := LoadSyncConfig(mergeOpts.configPath)
	if err != nil {
		return fmt.Errorf("加载配置文件失败: %v", err)
	}

	// 使用命令行参数覆盖配置文件设置
	inputDir := mergeOpts.inputDir
	if inputDir == "" {
		inputDir = config.Merge.InputDir
		if inputDir == "" {
			inputDir = config.Sync.OutputDir // 默认使用 sync 的输出目录
		}
	}

	outputDir := mergeOpts.outputDir
	if outputDir == "" {
		outputDir = config.Merge.OutputDir
	}

	filename := mergeOpts.filename
	if filename == "" {
		filename = config.Merge.Filename
	}

	headerTitle := config.Merge.HeaderTitle
	if headerTitle == "" {
		headerTitle = "合并的文档集合"
	}

	fmt.Printf("开始合并 %s 目录中的 .md 文件...\n", inputDir)
	fmt.Printf("配置来源: %s\n", getConfigSource(mergeOpts.configPath))

	// 检查输入目录是否存在
	if _, err := os.Stat(inputDir); os.IsNotExist(err) {
		return fmt.Errorf("输入目录不存在: %s", inputDir)
	}

	// 确保输出目录存在
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return fmt.Errorf("创建输出目录失败: %v", err)
	}

	// 查找所有 .md 文件
	mdFiles, err := findMarkdownFiles(inputDir)
	if err != nil {
		return fmt.Errorf("查找 .md 文件失败: %v", err)
	}
	if len(mdFiles) == 0 {
		return fmt.Errorf("在目录 %s 中未找到 .md 文件", inputDir)
	}

	// 按文件名排序（如果配置允许）
	if config.Merge.SortFiles {
		sort.Strings(mdFiles)
	}

	// 合并所有 Markdown 文件
	outputPath := filepath.Join(outputDir, filename)
	if err := mergeMarkdownFiles(mdFiles, outputPath, config.Merge, mergeOpts.original); err != nil {
		return fmt.Errorf("合并文件失败: %v", err)
	}

	fmt.Printf("✅ 成功合并 %d 个文件到: %s\n", len(mdFiles), outputPath)

	// 如果配置了 CSV 合并文件名（兼容 filename_csv 与 csv_filename），则另外生成一个仅合并 CSV 的 Markdown 文件
	csvOutName := config.Merge.FilenameCSV
	if strings.TrimSpace(csvOutName) == "" {
		csvOutName = config.Merge.CSVFilename
	}
	if strings.TrimSpace(csvOutName) != "" {
		// 查找 CSV 文件
		csvFiles, err := findCSVFiles(inputDir)
		if err != nil {
			return fmt.Errorf("查找 .csv 文件失败: %v", err)
		}
		if len(csvFiles) == 0 {
			fmt.Println("ℹ️ 未发现任何 CSV 文件，跳过 CSV 合并输出")
			return nil
		}
		if config.Merge.SortFiles {
			sort.Strings(csvFiles)
		}

		csvOut := filepath.Join(outputDir, csvOutName)
		if err := mergeCSVFilesToMarkdown(csvFiles, csvOut, config.Merge); err != nil {
			return fmt.Errorf("合并 CSV 为 Markdown 失败: %v", err)
		}
		fmt.Printf("✅ 成功合并 %d 个 CSV 到: %s\n", len(csvFiles), csvOut)
	}

	return nil
}

// getConfigSource 返回配置文件的来源描述
func getConfigSource(configPath string) string {
	if configPath != "" {
		return configPath
	}
	if _, err := os.Stat("config.yml"); err == nil {
		return "config.yml (当前目录)"
	}
	if _, err := os.Stat("config.yaml"); err == nil {
		return "config.yaml (当前目录)"
	}
	if _, err := os.Stat("sync_config.yaml"); err == nil {
		return "sync_config.yaml (当前目录，建议重命名为 config.yml)"
	}
	return "默认配置"
}

// findMarkdownFiles 查找指定目录中的所有 .md 文件
func findMarkdownFiles(dir string) ([]string, error) {
	var mdFiles []string

	err := filepath.WalkDir(dir, func(path string, d fs.DirEntry, err error) error {
		if err != nil {
			return err
		}

		if !d.IsDir() && strings.HasSuffix(strings.ToLower(d.Name()), ".md") {
			mdFiles = append(mdFiles, path)
		}

		return nil
	})

	return mdFiles, err
}

// findCSVFiles 查找指定目录中的所有 .csv 文件
func findCSVFiles(dir string) ([]string, error) {
	var csvFiles []string

	err := filepath.WalkDir(dir, func(path string, d fs.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if !d.IsDir() && strings.HasSuffix(strings.ToLower(d.Name()), ".csv") {
			csvFiles = append(csvFiles, path)
		}
		return nil
	})

	return csvFiles, err
}

// mergeMarkdownFiles 将多个 .md 文件合并为一个文件
func mergeMarkdownFiles(files []string, outputPath string, mergeConfig MergeSettings, original bool) error {
	outputFile, err := os.Create(outputPath)
	if err != nil {
		return err
	}
	defer outputFile.Close()

	// 写入文件头
	var header string
	if original {
		header = fmt.Sprintf(`# %s

> 此文件由 feishu2md 工具自动生成`, mergeConfig.HeaderTitle)

		if mergeConfig.IncludeTimestamp {
			header += fmt.Sprintf(`
> 生成时间: %s`, time.Now().Format("2006-01-02 15:04:05"))
		}

		header += fmt.Sprintf(`
> 包含文档数量: %d

---

`, len(files))
	} else {
		if mergeConfig.IncludeTimestamp {
			header = fmt.Sprintf(`> 生成时间: %s

`, time.Now().Format("2006-01-02 15:04:05"))
		} else {
			header = ""
		}
	}

	if _, err := outputFile.WriteString(header); err != nil {
		return err
	}

	// 逐个处理每个文件
	for i, filePath := range files {
		fmt.Printf("正在处理文件 (%d/%d): %s\n", i+1, len(files), filepath.Base(filePath))

		// 读取文件内容
		content, err := os.ReadFile(filePath)
		if err != nil {
			fmt.Printf("⚠️  读取文件失败，跳过: %s - %v\n", filePath, err)
			continue
		}

		// 获取文件名（不包含扩展名）作为大标题
		filename := strings.TrimSuffix(filepath.Base(filePath), ".md")

		// 写入分割线和大标题
		separator := fmt.Sprintf("\n\n---\n\n# 📄 %s\n\n", filename)
		if !original {
			separator = fmt.Sprintf("\n\n# 📄 %s\n\n", filename)
		}
		if _, err := outputFile.WriteString(separator); err != nil {
			return err
		}

		// 写入文件内容
		contentStr := string(content)

		// 如果文件以 # 开头，将其转换为 ## 以避免与我们的大标题冲突
		lines := strings.Split(contentStr, "\n")
		for j, line := range lines {
			if strings.HasPrefix(strings.TrimSpace(line), "#") {
				// 在现有的 # 前面再加一个 #
				lines[j] = "#" + line
			}
		}
		contentStr = strings.Join(lines, "\n")

		writeStr := contentStr
		if !original {
			writeStr = compactMarkdown(contentStr, mergeConfig)
		}
		if _, err := outputFile.WriteString(writeStr); err != nil {
			return err
		}

		// 确保文件内容后有换行
		if !strings.HasSuffix(contentStr, "\n") {
			if _, err := outputFile.WriteString("\n"); err != nil {
				return err
			}
		}
	}

	// 写入文件尾
	footer := fmt.Sprintf("\n\n---\n\n> 文档合并完成 | 总计 %d 个文件", len(files))

	if mergeConfig.IncludeTimestamp {
		footer += fmt.Sprintf(" | 生成时间: %s", time.Now().Format("2006-01-02 15:04:05"))
	}

	footer += "\n"

	if original {
		if _, err := outputFile.WriteString(footer); err != nil {
			return err
		}
	}

	return nil
}

// mergeCSVFilesToMarkdown 将多个 CSV 文件合并为一个 Markdown（以标题 + 原始CSV代码块形式展示）
func mergeCSVFilesToMarkdown(files []string, outputPath string, mergeConfig MergeSettings) error {
	out, err := os.Create(outputPath)
	if err != nil {
		return err
	}
	defer out.Close()

	// 头部（仅时间戳）
	if mergeConfig.IncludeTimestamp {
		if _, err := out.WriteString(fmt.Sprintf("> 生成时间: %s\n\n", time.Now().Format("2006-01-02 15:04:05"))); err != nil {
			return err
		}
	}

	for i, filePath := range files {
		fmt.Printf("正在处理 CSV (%d/%d): %s\n", i+1, len(files), filepath.Base(filePath))
		title := strings.TrimSuffix(filepath.Base(filePath), filepath.Ext(filePath))
		// 大标题
		if _, err := out.WriteString(fmt.Sprintf("# %s\n\n", title)); err != nil {
			return err
		}

		// 读取 CSV 原文并去除 UTF-8 BOM
		raw, rerr := os.ReadFile(filePath)
		if rerr != nil {
			return rerr
		}
		if len(raw) >= 3 && raw[0] == 0xEF && raw[1] == 0xBB && raw[2] == 0xBF {
			raw = raw[3:]
		}
		if _, err := out.WriteString("```csv\n"); err != nil {
			return err
		}
		if _, err := out.Write(raw); err != nil {
			return err
		}
		// 确保以换行结尾，再关闭代码块
		if len(raw) == 0 || raw[len(raw)-1] != '\n' {
			if _, err := out.WriteString("\n"); err != nil {
				return err
			}
		}
		if _, err := out.WriteString("```\n\n"); err != nil {
			return err
		}
	}

	return nil
}

// 保持代码块不变；移除 HR；图片转 [img]；链接转 文本 [url]；裸 URL -> [url]；压缩标准表格
func compactMarkdown(input string, mergeConfig MergeSettings) string {
	lines := strings.Split(input, "\n")
	var out []string
	inCode := false
	fence := ""

	i := 0
	for i < len(lines) {
		line := lines[i]
		trimmed := strings.TrimSpace(line)

		// HTML 表格压缩：检测 <table> ... </table>
		if !inCode && strings.Contains(strings.ToLower(trimmed), "<table") {
			// 收集整个表格块
			start := i
			j := i
			foundEnd := false
			for j < len(lines) {
				if strings.Contains(strings.ToLower(lines[j]), "</table>") {
					foundEnd = true
					break
				}
				j++
			}
			if foundEnd {
				tableBlock := strings.Join(lines[start:j+1], "\n")
				dict := compressHTMLTableBlock(tableBlock, mergeConfig)
				if dict != "" {
					out = append(out, dict)
					i = j + 1
					continue
				}
				// 如果无法压缩，原样输出
				out = append(out, lines[start:j+1]...)
				i = j + 1
				continue
			}
			// 未找到闭合标签，则继续常规处理
		}

		// 代码围栏
		if strings.HasPrefix(trimmed, "```") || strings.HasPrefix(trimmed, "~~~") {
			mark := trimmed[:3]
			if !inCode {
				inCode = true
				fence = mark
			} else if strings.HasPrefix(trimmed, fence) {
				inCode = false
				fence = ""
			}
			out = append(out, line)
			i++
			continue
		}

		if inCode {
			out = append(out, line)
			i++
			continue
		}

		// HR 移除
		if isHRLine(trimmed) {
			i++
			continue
		}

		// 表格压缩
		if looksLikeTableHeader(line) && i+1 < len(lines) && isTableDelimiter(lines[i+1]) {
			i += 2 // 跳过表头与分隔
			for i < len(lines) && isTableRow(lines[i]) {
				out = append(out, compressTableRow(lines[i]))
				i++
			}
			continue
		}

		processed := simplifyLine(line)
		if processed != "" {
			out = append(out, processed)
		}
		i++
	}

	return strings.Join(out, "\n")
}

func isHRLine(s string) bool {
	t := strings.TrimSpace(s)
	return t == "---" || t == "***" || t == "___"
}

func looksLikeTableHeader(s string) bool {
	return strings.Contains(s, "|") && !isTableDelimiter(s)
}

func isTableDelimiter(s string) bool {
	t := strings.TrimSpace(s)
	if !strings.Contains(t, "|") && !strings.Contains(t, "-") {
		return false
	}
	for _, ch := range t {
		if ch != '|' && ch != '-' && ch != ':' && ch != ' ' && ch != '\t' {
			return false
		}
	}
	return strings.Contains(t, "---")
}

func isTableRow(s string) bool {
	return strings.Contains(s, "|")
}

func compressTableRow(row string) string {
	r := strings.TrimSpace(row)
	r = strings.TrimPrefix(r, "|")
	r = strings.TrimSuffix(r, "|")
	parts := strings.Split(r, "|")
	cells := make([]string, 0, len(parts))
	for _, p := range parts {
		c := strings.TrimSpace(p)
		if strings.HasPrefix(c, "`") && strings.HasSuffix(c, "`") && len(c) >= 2 {
			c = strings.TrimSuffix(strings.TrimPrefix(c, "`"), "`")
		}
		cells = append(cells, c)
	}
	return strings.Join(cells, ":")
}

func simplifyLine(line string) string {
	s := line
	s = replaceImagesWithMarker(s)
	s = replaceLinksKeepText(s)
	s = replaceBareURL(s)
	// 移除 blockquote 的工具说明与数量行
	trimmed := strings.TrimSpace(s)
	if strings.HasPrefix(trimmed, ">") {
		t := strings.TrimSpace(strings.TrimPrefix(trimmed, ">"))
		if strings.Contains(t, "feishu2md") || strings.Contains(t, "此文件由") || strings.Contains(t, "包含文档数量") {
			return ""
		}
	}
	return s
}

func replaceImagesWithMarker(s string) string {
	for {
		start := strings.Index(s, "![")
		if start == -1 {
			break
		}
		mid := strings.Index(s[start:], "](")
		if mid == -1 {
			break
		}
		mid = start + mid
		end := strings.Index(s[mid+2:], ")")
		if end == -1 {
			break
		}
		end = mid + 2 + end
		s = s[:start] + "[img]" + s[end+1:]
	}
	return s
}

func replaceLinksKeepText(s string) string {
	for {
		open := strings.Index(s, "[")
		if open == -1 {
			break
		}
		close := strings.Index(s[open:], "](")
		if close == -1 {
			break
		}
		close = open + close
		end := strings.Index(s[close+2:], ")")
		if end == -1 {
			break
		}
		end = close + 2 + end
		text := s[open+1 : close]
		s = s[:open] + text + " [url]" + s[end+1:]
	}
	return s
}

func replaceBareURL(s string) string {
	for {
		idx := strings.Index(s, "http://")
		idxs := strings.Index(s, "https://")
		if idx == -1 || (idxs != -1 && idxs < idx) {
			idx = idxs
		}
		if idx == -1 {
			break
		}
		end := idx
		for end < len(s) {
			ch := s[end]
			if ch == ' ' || ch == ')' || ch == ']' || ch == '"' || ch == '\n' {
				break
			}
			end++
		}
		s = s[:idx] + "[url]" + s[end:]
	}
	return s
}

// ---------- HTML Table compaction ----------
func compressHTMLTableBlock(tableHTML string, mergeConfig MergeSettings) string {
	// Extract <tr> rows
	trRe := regexp.MustCompile(`(?is)<tr[^>]*>(.*?)</tr>`)
	tdRe := regexp.MustCompile(`(?is)<td[^>]*>(.*?)</td>`)
	brRe := regexp.MustCompile(`(?is)<br\s*/?>`)
	tagRe := regexp.MustCompile(`(?is)<[^>]+>`) // strip any remaining tags

	// Helper to clean cell text
	clean := func(s string) string {
		s = brRe.ReplaceAllString(s, " ")
		s = tagRe.ReplaceAllString(s, "")
		s = html.UnescapeString(s)
		s = strings.ReplaceAll(s, "**", "")
		s = strings.TrimSpace(s)
		// trim outer backticks
		if strings.HasPrefix(s, "`") && strings.HasSuffix(s, "`") && len(s) >= 2 {
			s = strings.TrimSuffix(strings.TrimPrefix(s, "`"), "`")
		}
		return s
	}

	// Build table: slice of rows
	var rows [][]string
	for _, m := range trRe.FindAllStringSubmatch(tableHTML, -1) {
		inner := m[1]
		var cells []string
		for _, c := range tdRe.FindAllStringSubmatch(inner, -1) {
			cells = append(cells, clean(c[1]))
		}
		// skip empty rows
		nonEmpty := false
		for _, c := range cells {
			if strings.TrimSpace(c) != "" {
				nonEmpty = true
				break
			}
		}
		if nonEmpty {
			rows = append(rows, cells)
		}
	}

	if len(rows) == 0 {
		return ""
	}

	// Detect header keywords to decide grouping strategy
	hasHeader := false
	headerKeys := mergeConfig.GroupHeaderKeywords
	if len(rows) > 0 {
		headerJoined := strings.Join(rows[0], " ")
		cnt := 0
		for _, k := range headerKeys {
			if strings.Contains(headerJoined, k) {
				cnt++
			}
		}
		if cnt >= 2 {
			hasHeader = true
		}
	}

	// If looks like category table with 3-4 cols, group by first col
	if hasHeader {
		groupOrder := []string{}
		itemsByGroup := map[string][]string{}
		currentGroup := ""

		for idx, row := range rows {
			// skip header row
			if idx == 0 {
				continue
			}
			// Identify group/code/name by column count
			g, code, cn := "", "", ""
			if len(row) >= 4 {
				g, code, cn = row[0], row[1], row[2]
			} else if len(row) == 3 {
				// likely no group cell due to rowspan
				g, code, cn = "", row[0], row[1]
			} else if len(row) == 2 {
				g, code = "", row[0]
				cn = row[1]
			} else {
				continue
			}

			if strings.TrimSpace(g) != "" {
				currentGroup = g
				if _, ok := itemsByGroup[currentGroup]; !ok {
					groupOrder = append(groupOrder, currentGroup)
					itemsByGroup[currentGroup] = []string{}
				}
			}

			if currentGroup == "" {
				// can't place without a group
				continue
			}

			code = strings.TrimSpace(code)
			cn = strings.TrimSpace(cn)
			if code == "" {
				continue
			}
			item := code
			if cn != "" {
				item = fmt.Sprintf("%s(%s)", code, cn)
			}
			itemsByGroup[currentGroup] = append(itemsByGroup[currentGroup], item)
		}

		// If no groups collected, fall back to generic
		if len(itemsByGroup) == 0 {
			return genericHTMLTableToLines(rows, mergeConfig)
		}

		var b strings.Builder
		for idx, g := range groupOrder {
			it := itemsByGroup[g]
			if len(it) == 0 {
				continue
			}
			if idx > 0 {
				b.WriteString("\n")
			}
			b.WriteString(fmt.Sprintf("%s: %s", g, strings.Join(it, ", ")))
		}
		return b.String()
	}

	// Fallback: generic colon-joined rows
	return genericHTMLTableToLines(rows, mergeConfig)
}

func genericHTMLTableToLines(rows [][]string, mergeConfig MergeSettings) string {
	if len(rows) == 0 {
		return ""
	}
	// Try to detect header row and skip
	start := 0
	if looksHeaderRow(rows[0], mergeConfig.HeaderKeywords) {
		start = 1
	}
	var out []string
	for i := start; i < len(rows); i++ {
		cells := rows[i]
		vals := make([]string, 0, len(cells))
		for _, c := range cells {
			if c == "[img]" || c == "img" {
				continue
			}
			vals = append(vals, strings.TrimSpace(c))
		}
		if len(vals) == 0 {
			continue
		}
		out = append(out, strings.Join(vals, ":"))
	}
	return strings.Join(out, "\n")
}

func looksHeaderRow(cells []string, keywords []string) bool {
	if len(cells) == 0 {
		return false
	}
	joined := strings.Join(cells, " ")
	keys := keywords
	if len(keys) == 0 {
		return false
	}
	hits := 0
	for _, k := range keys {
		if strings.Contains(joined, k) {
			hits++
		}
	}
	return hits >= 1
}
