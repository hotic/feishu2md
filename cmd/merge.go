package main

import (
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
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

	// 合并文件
	outputPath := filepath.Join(outputDir, filename)
	if err := mergeMarkdownFiles(mdFiles, outputPath, config.Merge); err != nil {
		return fmt.Errorf("合并文件失败: %v", err)
	}

	fmt.Printf("✅ 成功合并 %d 个文件到: %s\n", len(mdFiles), outputPath)
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

// mergeMarkdownFiles 将多个 .md 文件合并为一个文件
func mergeMarkdownFiles(files []string, outputPath string, mergeConfig MergeSettings) error {
	outputFile, err := os.Create(outputPath)
	if err != nil {
		return err
	}
	defer outputFile.Close()

	// 写入文件头
	header := fmt.Sprintf(`# %s

> 此文件由 feishu2md 工具自动生成`, mergeConfig.HeaderTitle)

	if mergeConfig.IncludeTimestamp {
		header += fmt.Sprintf(`
> 生成时间: %s`, time.Now().Format("2006-01-02 15:04:05"))
	}

	header += fmt.Sprintf(`
> 包含文档数量: %d

---

`, len(files))

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

		if _, err := outputFile.WriteString(contentStr); err != nil {
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

	if _, err := outputFile.WriteString(footer); err != nil {
		return err
	}

	return nil
}
