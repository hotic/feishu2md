package main

import (
	"context"
	"crypto/sha256"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/Wsine/feishu2md/core"
	"github.com/Wsine/feishu2md/utils"
	"github.com/chyroc/lark"
	"github.com/urfave/cli/v2"
)

type SyncOpts struct {
	configPath string
	group      string
	force      bool
}

var syncOpts = SyncOpts{}

// getSyncCommand returns the sync command definition
func getSyncCommand() *cli.Command {
	return &cli.Command{
		Name:  "sync",
		Usage: "Manage and sync multiple documents via configuration file",
		Subcommands: []*cli.Command{
			{
				Name:  "init",
				Usage: "Initialize sync configuration file",
				Flags: []cli.Flag{
					&cli.StringFlag{
						Name:        "config",
						Usage:       "Path to config file",
						Destination: &syncOpts.configPath,
					},
				},
				Action: handleSyncInit,
			},
			{
				Name:      "add",
				Usage:     "Add a document to sync configuration",
				ArgsUsage: "<url>",
				Flags: []cli.Flag{
					&cli.StringFlag{
						Name:     "name",
						Usage:    "Document name",
						Required: true,
					},
					&cli.StringFlag{
						Name:  "group",
						Usage: "Document group",
						Value: "default",
					},
					&cli.StringFlag{
						Name:        "config",
						Usage:       "Path to config file",
						Destination: &syncOpts.configPath,
					},
				},
				Action: handleSyncAdd,
			},
			{
				Name:  "list",
				Usage: "List all configured documents",
				Flags: []cli.Flag{
					&cli.StringFlag{
						Name:        "group",
						Usage:       "Filter by group",
						Destination: &syncOpts.group,
					},
					&cli.StringFlag{
						Name:        "config",
						Usage:       "Path to config file",
						Destination: &syncOpts.configPath,
					},
				},
				Action: handleSyncList,
			},
			{
				Name:  "run",
				Usage: "Run sync for configured documents",
				Flags: []cli.Flag{
					&cli.StringFlag{
						Name:        "group",
						Usage:       "Sync only specific group",
						Destination: &syncOpts.group,
					},
					&cli.BoolFlag{
						Name:        "force",
						Usage:       "Force re-download all documents",
						Destination: &syncOpts.force,
					},
					&cli.StringFlag{
						Name:        "config",
						Usage:       "Path to config file",
						Destination: &syncOpts.configPath,
					},
				},
				Action: handleSyncRun,
			},
			{
				Name:      "remove",
				Usage:     "Remove a document from configuration",
				ArgsUsage: "<name or index>",
				Flags: []cli.Flag{
					&cli.StringFlag{
						Name:        "config",
						Usage:       "Path to config file",
						Destination: &syncOpts.configPath,
					},
				},
				Action: handleSyncRemove,
			},
		},
	}
}

func handleSyncInit(ctx *cli.Context) error {
	config := NewSyncConfig()

	// Add example document
	config.Documents = []DocConfig{
		{
			Name:  "示例文档",
			URL:   "https://example.feishu.cn/docx/example",
			Group: "示例",
		},
	}

	if err := config.Save(syncOpts.configPath); err != nil {
		return fmt.Errorf("failed to save config: %v", err)
	}

	configPath := syncOpts.configPath
	if configPath == "" {
		configPath, _ = GetSyncConfigPath()
	}

	fmt.Printf("Sync configuration initialized at: %s\n", configPath)
	fmt.Println("Please edit the configuration file to add your documents.")
	fmt.Println("\nNext steps:")
	fmt.Println("1. Configure app credentials: feishu2md config --appId <id> --appSecret <secret>")
	fmt.Println("2. Add documents: feishu2md sync add <url> --name <name> --group <group>")
	fmt.Println("3. Run sync: feishu2md sync run")

	return nil
}

func handleSyncAdd(ctx *cli.Context) error {
	if ctx.NArg() == 0 {
		return cli.Exit("Please specify document URL", 1)
	}

	url := ctx.Args().First()
	name := ctx.String("name")
	group := ctx.String("group")

	config, err := LoadSyncConfig(syncOpts.configPath)
	if err != nil {
		return fmt.Errorf("failed to load config: %v", err)
	}

	if err := config.AddDocument(name, url, group); err != nil {
		return err
	}

	if err := config.Save(syncOpts.configPath); err != nil {
		return fmt.Errorf("failed to save config: %v", err)
	}

	fmt.Printf("Added document '%s' to sync configuration\n", name)
	return nil
}

func handleSyncList(ctx *cli.Context) error {
	config, err := LoadSyncConfig(syncOpts.configPath)
	if err != nil {
		return fmt.Errorf("failed to load config: %v", err)
	}

	if len(config.Documents) == 0 {
		fmt.Println("No documents configured")
		fmt.Println("Use 'feishu2md sync add <url> --name <name>' to add documents")
		return nil
	}

	fmt.Println("\n=== Sync Configuration ===")
	fmt.Printf("Output Directory: %s\n", config.Sync.OutputDir)
	fmt.Printf("Organize by Group: %v\n", config.Sync.OrganizeByGroup)
	fmt.Printf("Concurrent Downloads: %d\n", config.Sync.ConcurrentDownloads)
	fmt.Printf("Sync Mode: %s\n", config.Sync.SyncMode)

	fmt.Println("\n=== Configured Documents ===")

	// Group documents by group
	groups := make(map[string][]DocConfig)
	for _, doc := range config.Documents {
		if syncOpts.group != "" && doc.Group != syncOpts.group {
			continue
		}
		groups[doc.Group] = append(groups[doc.Group], doc)
	}

	if len(groups) == 0 {
		fmt.Printf("No documents found for group '%s'\n", syncOpts.group)
		return nil
	}

	index := 0
	for group, docs := range groups {
		groupName := group
		if groupName == "" {
			groupName = "根目录"
		}
		fmt.Printf("\n[%s]\n", groupName)
		for _, doc := range docs {
			// Determine type: explicit config overrides URL auto-detect
			docType := doc.Type
			if docType == "" {
				docType = "docx"
				if strings.Contains(doc.URL, "/wiki/settings/") {
					docType = "wiki_space"
				} else if strings.Contains(doc.URL, "/wiki/") {
					docType = "wiki_page"
				} else if strings.Contains(doc.URL, "/drive/folder/") || strings.Contains(doc.URL, "/folder/") {
					docType = "folder"
				}
			}
			fmt.Printf("  %d. %s (%s)\n", index, doc.Name, docType)
			fmt.Printf("     URL: %s\n", doc.URL)
			index++
		}
	}

	fmt.Printf("\nTotal: %d documents\n", index)
	return nil
}

func handleSyncRun(ctx *cli.Context) error {
	// Load sync configuration
	syncConfig, err := LoadSyncConfig(syncOpts.configPath)
	if err != nil {
		return fmt.Errorf("failed to load sync config: %v", err)
	}

	// Load feishu configuration
	configPath, err := core.GetConfigFilePath()
	if err != nil {
		return err
	}
	feishuConfig, err := core.ReadConfigFromFile(configPath)
	if err != nil {
		return fmt.Errorf("failed to load feishu config: %v\nPlease run 'feishu2md config --appId <id> --appSecret <secret>' first", err)
	}

	// Get documents to sync
	documents := syncConfig.GetDocuments(syncOpts.group)
	if len(documents) == 0 {
		fmt.Println("No documents to sync")
		fmt.Println("Please add documents to your configuration file")
		return nil
	}

	fmt.Printf("\nStarting sync for %d documents...\n", len(documents))
	fmt.Printf("Output directory: %s\n", syncConfig.Sync.OutputDir)
	fmt.Printf("Sync mode: %s\n", syncConfig.Sync.SyncMode)

	// 根据同步模式决定是否清理目录
	// clean_all: 总是清理
	// incremental: 不清理，但 --force 标志可以强制清理
	if syncConfig.Sync.SyncMode == "clean_all" || syncOpts.force {
		fmt.Println("Cleaning output directory...")
		if err := cleanOutputDirectory(syncConfig.Sync.OutputDir); err != nil {
			fmt.Printf("Warning: failed to clean output directory: %v\n", err)
		}
	}

	// Create client
	client := core.NewClient(feishuConfig.Feishu.AppId, feishuConfig.Feishu.AppSecret)
	ctx2 := context.Background()

	// 过滤需要同步的文档（增量模式）
	documentsToSync, err := filterDocumentsForSync(ctx2, client, documents, syncConfig.Sync.OutputDir, &syncConfig.Sync)
	if err != nil {
		return fmt.Errorf("failed to filter documents: %v", err)
	}

	if len(documentsToSync) == 0 {
		fmt.Println("No documents need to be synced")
		return nil
	}

	if len(documentsToSync) < len(documents) {
		fmt.Printf("Filtered %d documents, %d will be synced\n", len(documents), len(documentsToSync))
	}

	// Sync documents with concurrency control
	var wg sync.WaitGroup
	semaphore := make(chan struct{}, syncConfig.Sync.ConcurrentDownloads)
	errors := make([]error, 0)
	var errorsMux sync.Mutex

	startTime := time.Now()
	successCount := 0
	var successMux sync.Mutex

	for _, doc := range documentsToSync {
		wg.Add(1)
		semaphore <- struct{}{} // Acquire semaphore

		go func(doc DocConfig) {
			defer wg.Done()
			defer func() { <-semaphore }() // Release semaphore

			groupInfo := doc.Group
			if groupInfo == "" {
				groupInfo = "根目录"
			}
			fmt.Printf("\n[%s] 下载 %s...\n", groupInfo, doc.Name)

			outputDir := syncConfig.Sync.OutputDir
			// 只有当 OrganizeByGroup 为 true 且 group 不为空时才按组存储
			if syncConfig.Sync.OrganizeByGroup && doc.Group != "" {
				outputDir = filepath.Join(outputDir, doc.Group)
			}

			err := syncDocument(ctx2, client, doc, outputDir, feishuConfig, &syncConfig.Sync)
			if err != nil {
				errorsMux.Lock()
				errors = append(errors, fmt.Errorf("%s: %v", doc.Name, err))
				errorsMux.Unlock()
				fmt.Printf("  ✗ Failed: %v\n", err)
			} else {
				successMux.Lock()
				successCount++
				successMux.Unlock()
				fmt.Printf("  ✓ 成功: %s\n", doc.Name)
			}
		}(doc)
	}

	wg.Wait()

	// Print summary
	elapsed := time.Since(startTime)
	fmt.Printf("\n=== 同步完成 ===\n")
	fmt.Printf("耗时: %v\n", elapsed.Round(time.Second))
	fmt.Printf("成功: %d/%d\n", successCount, len(documentsToSync))

	if len(errors) > 0 {
		fmt.Println("\n错误:")
		for _, err := range errors {
			fmt.Printf("  - %v\n", err)
		}
		return cli.Exit("同步完成但有错误", 1)
	}

	fmt.Println("\n✓ 所有文档同步成功!")
	return nil
}

func handleSyncRemove(ctx *cli.Context) error {
	if ctx.NArg() == 0 {
		return cli.Exit("Please specify document name or index to remove", 1)
	}

	nameOrIndex := ctx.Args().First()

	config, err := LoadSyncConfig(syncOpts.configPath)
	if err != nil {
		return fmt.Errorf("failed to load config: %v", err)
	}

	if err := config.RemoveDocument(nameOrIndex); err != nil {
		return err
	}

	if err := config.Save(syncOpts.configPath); err != nil {
		return fmt.Errorf("failed to save config: %v", err)
	}

	fmt.Printf("Removed document '%s' from configuration\n", nameOrIndex)
	return nil
}

// syncDocument syncs a single document based on its type
func syncDocument(ctx context.Context, client *core.Client, doc DocConfig, outputDir string, config *core.Config, syncSettings *SyncSettings) error {
	dlConfig = *config // Set global dlConfig

	// Create output directory if it doesn't exist
	if _, err := os.Stat(outputDir); os.IsNotExist(err) {
		if err := os.MkdirAll(outputDir, 0755); err != nil {
			return fmt.Errorf("failed to create output directory: %v", err)
		}
	}

	// Determine type: explicit config overrides URL auto-detect
	docType := doc.Type
	if docType == "" {
		docType = "docx"
		if strings.Contains(doc.URL, "/wiki/") {
			docType = "wiki"
		} else if strings.Contains(doc.URL, "/folder/") {
			docType = "folder"
		}
	}

	// 判断是否跳过图片下载：单文档配置优先级高于全局配置
	skipImages := syncSettings.SkipImages // 默认使用全局配置
	if doc.SkipImages != nil {
		skipImages = *doc.SkipImages // 如果单文档有设置，则使用单文档配置
	}

	// 决定是否使用原始标题名
	var docName string
	if syncSettings.UseOriginalTitle {
		docName = "" // 空字符串表示使用原始标题
	} else {
		docName = doc.Name // 使用配置中的自定义名称
	}

	opts := DownloadOpts{
		outputDir:        outputDir,
		dump:             false,
		batch:            docType == "folder",
		wiki:             docType == "wiki_space",
		docName:          docName, // 根据配置决定使用哪个名称
		skipImages:       skipImages,
		useOriginalTitle: syncSettings.UseOriginalTitle, // 传递新的配置选项
	}

	switch docType {
	case "wiki_space":
		// For entire wiki spaces (with /wiki/settings/ URL)
		return downloadWiki(ctx, client, doc.URL)
	case "wiki_page":
		// For individual wiki pages, treat them as documents
		actualFileName, err := downloadDocument(ctx, client, doc.URL, &opts)
		if err != nil {
			return err
		}
		// 下载成功后，保存元数据（用于增量同步）
		if syncSettings.SyncMode == "incremental" {
			return saveDocumentMetadataWithFileName(ctx, client, doc, outputDir, syncSettings, actualFileName)
		}
		return nil
	case "folder":
		return downloadDocuments(ctx, client, doc.URL)
	case "csv":
		actualFileName, err := exportBitable(ctx, client, doc.URL, "csv", outputDir, docName)
		if err != nil {
			return err
		}
		// 下载成功后，保存元数据（用于增量同步）
		if syncSettings.SyncMode == "incremental" {
			return saveDocumentMetadataWithFileName(ctx, client, doc, outputDir, syncSettings, actualFileName)
		}
		return nil
	case "xlsx":
		actualFileName, err := exportBitable(ctx, client, doc.URL, "xlsx", outputDir, docName)
		if err != nil {
			return err
		}
		// 下载成功后，保存元数据（用于增量同步）
		if syncSettings.SyncMode == "incremental" {
			return saveDocumentMetadataWithFileName(ctx, client, doc, outputDir, syncSettings, actualFileName)
		}
		return nil
	default: // docx
		// 先执行下载
		actualFileName, err := downloadDocument(ctx, client, doc.URL, &opts)
		if err != nil {
			return err
		}

		// 下载成功后，保存元数据（用于增量同步）
		if syncSettings.SyncMode == "incremental" {
			return saveDocumentMetadataWithFileName(ctx, client, doc, outputDir, syncSettings, actualFileName)
		}
		return nil
	}
}

// cleanOutputDirectory removes all files in the output directory
func cleanOutputDirectory(dir string) error {
	if dir == "" || dir == "/" || dir == "." {
		return fmt.Errorf("invalid directory path")
	}

	// Check if directory exists
	if _, err := os.Stat(dir); os.IsNotExist(err) {
		return nil // Directory doesn't exist, nothing to clean
	}

	// Read directory contents
	entries, err := os.ReadDir(dir)
	if err != nil {
		return err
	}

	// Remove each entry
	for _, entry := range entries {
		path := filepath.Join(dir, entry.Name())
		if err := os.RemoveAll(path); err != nil {
			return fmt.Errorf("failed to remove %s: %v", path, err)
		}
	}

	return nil
}

// 检查是否需要同步某个文档（用于增量模式）
func shouldSyncDocument(ctx context.Context, client *core.Client, doc DocConfig, outputDir string, syncSettings *SyncSettings) (bool, error) {
	if syncSettings.SyncMode != "incremental" {
		return true, nil // 非增量模式，总是同步
	}

	actualOutputDir := outputDir
	if syncSettings.OrganizeByGroup && doc.Group != "" {
		actualOutputDir = filepath.Join(outputDir, doc.Group)
	}

	// 特殊处理xlsx/csv类型（优先处理，避免被UseOriginalTitle逻辑影响）
	if doc.Type == "xlsx" || doc.Type == "csv" {
		// 对于表格文件，文件名由系统生成，需要检查目录
		result, err := checkTableDocumentExists(actualOutputDir, doc.URL, doc.Type)
		return result, err
	}

	// 首先获取文档信息以确定实际的文件名
	docType, docToken, err := utils.ValidateDocumentURL(doc.URL)
	if err != nil {
		return true, nil // URL解析失败，假设需要更新
	}

	var fileName string
	var metadataPath string

	// 创建元数据目录
	metadataDir := filepath.Join(actualOutputDir, ".feishu2md")

	if syncSettings.UseOriginalTitle && docType == "docx" {
		// 使用原始标题的情况，需要先获取文档标题
		docx, _, err := client.GetDocxContent(ctx, docToken)
		if err != nil {
			return true, nil // 获取失败，假设需要更新
		}
		fileName = fmt.Sprintf("%s.md", utils.SanitizeFileName(docx.Title))
		metadataPath = filepath.Join(metadataDir, fmt.Sprintf("%s.meta", utils.SanitizeFileName(docx.Title)))
	} else if syncSettings.UseOriginalTitle {
		// 非docx文档使用原始标题的情况，暂时无法预测文件名，需要检查目录中的文件
		return checkDocumentByURL(actualOutputDir, doc.URL)
	} else {
		// 使用配置中的名称
		fileName = fmt.Sprintf("%s.md", utils.SanitizeFileName(doc.Name))
		metadataPath = filepath.Join(metadataDir, fmt.Sprintf("%s.meta", utils.SanitizeFileName(doc.Name)))
	}

	filePath := filepath.Join(actualOutputDir, fileName)

	// 检查文件是否存在
	_, err = os.Stat(filePath)
	if os.IsNotExist(err) {
		// 文件不存在，需要下载
		return true, nil
	}
	if err != nil {
		return false, err
	}

	// 文件存在，检查是否有元数据文件和版本信息
	metadataData, err := os.ReadFile(metadataPath)
	if err != nil {
		// 没有元数据文件，假设需要更新
		return true, nil
	}

	// 从元数据中获取上次同步的RevisionID
	lines := strings.Split(string(metadataData), "\n")
	var lastRevisionID string
	for _, line := range lines {
		if strings.HasPrefix(line, "RevisionID=") {
			lastRevisionID = strings.TrimPrefix(line, "RevisionID=")
			break
		}
	}

	if lastRevisionID == "" {
		// 没有找到RevisionID，需要更新
		return true, nil
	}

	// 获取当前文档的RevisionID来比较（只对docx有效）
	if docType == "docx" {
		var currentDocx *lark.DocxDocument
		var currentRevisionID string

		if syncSettings.UseOriginalTitle {
			// 如果已经获取过文档（为了得到文件名），就重用结果
			// 否则重新获取
			currentDocx, _, err = client.GetDocxContent(ctx, docToken)
			if err != nil {
				// 获取失败，假设需要更新
				return true, nil
			}
		} else {
			currentDocx, _, err = client.GetDocxContent(ctx, docToken)
			if err != nil {
				// 获取失败，假设需要更新
				return true, nil
			}
		}

		// 比较RevisionID
		currentRevisionID = fmt.Sprintf("%d", currentDocx.RevisionID)
		if currentRevisionID != lastRevisionID {
			fmt.Printf("检测到文档 %s 有更新 (RevisionID: %s -> %s)\n", doc.Name, lastRevisionID, currentRevisionID)
			return true, nil
		}

		// RevisionID相同，跳过
		return false, nil
	}

	// 非docx文档（如wiki），通过内容哈希检测更新
	// 获取当前文档内容的哈希值来比较
	
	// 从元数据中获取上次保存的内容哈希
	var lastContentHash string
	for _, line := range lines {
		if strings.HasPrefix(line, "ContentHash=") {
			lastContentHash = strings.TrimPrefix(line, "ContentHash=")
			break
		}
	}
	
	// 获取当前文档内容并计算哈希
	currentDocx, currentBlocks, err := client.GetDocxContent(ctx, docToken)
	if err != nil {
		// 获取失败，保守起见重新同步
		fmt.Printf("获取文档 %s 内容失败，重新同步: %v\n", doc.Name, err)
		return true, nil
	}
	
	// 计算当前内容的哈希值（使用文档标题+内容块）
	parser := core.NewParser(core.OutputConfig{})
	currentContent := parser.ParseDocxContent(currentDocx, currentBlocks)
	currentContentHash := fmt.Sprintf("%x", sha256.Sum256([]byte(currentDocx.Title+currentContent)))
	
	if lastContentHash == "" {
		// 没有找到内容哈希，需要更新
		fmt.Printf("文档 %s 没有内容哈希记录，重新同步\n", doc.Name)
		return true, nil
	}
	
	if currentContentHash != lastContentHash {
		fmt.Printf("检测到文档 %s 内容有更新\n", doc.Name)
		return true, nil
	}
	
	// 内容哈希相同，跳过
	return false, nil
}

// 通过URL检查文档是否存在（用于无法预测文件名的情况）
func checkDocumentByURL(outputDir, url string) (bool, error) {
	// 查找元数据目录中是否有与此URL相关的元数据文件
	metadataDir := filepath.Join(outputDir, ".feishu2md")
	entries, err := os.ReadDir(metadataDir)
	if err != nil {
		return true, nil // 元数据目录不存在，需要下载
	}

	for _, entry := range entries {
		if strings.HasSuffix(entry.Name(), ".meta") {
			metadataPath := filepath.Join(metadataDir, entry.Name())
			data, err := os.ReadFile(metadataPath)
			if err != nil {
				continue
			}

			lines := strings.Split(string(data), "\n")
			var storedURL, documentName, actualFileName string
			for _, line := range lines {
				if strings.HasPrefix(line, "URL=") {
					storedURL = strings.TrimPrefix(line, "URL=")
				} else if strings.HasPrefix(line, "DocumentName=") {
					// 保持向后兼容，但优先使用ActualFileName
					documentName = strings.TrimPrefix(line, "DocumentName=")
				} else if strings.HasPrefix(line, "ActualFileName=") {
					actualFileName = strings.TrimPrefix(line, "ActualFileName=")
				}
			}

			if storedURL == url {
				// 找到对应的元数据，优先使用实际文件名
				var filePath string
				if actualFileName != "" {
					filePath = filepath.Join(outputDir, actualFileName)
				} else if documentName != "" {
					// 向后兼容：如果没有ActualFileName，使用DocumentName
					filePath = filepath.Join(outputDir, fmt.Sprintf("%s.md", documentName))
				} else {
					// 没有文件名信息，需要重新下载
					return true, nil
				}

				_, err := os.Stat(filePath)
				if err == nil {
					// 文档文件存在，不需要下载
					return false, nil
				}
				// 元数据存在但文档文件不存在，需要重新下载
				return true, nil
			}
		}
	}

	// 没有找到对应的文档，需要下载
	return true, nil
}

// 检查表格文档是否存在
func checkTableDocumentExists(outputDir, url, docType string) (bool, error) {
	// 查找元数据目录中是否有与此URL相关的元数据文件
	metadataDir := filepath.Join(outputDir, ".feishu2md")
	entries, err := os.ReadDir(metadataDir)
	if err != nil {
		return true, nil // 元数据目录不存在，需要下载
	}

	var actualFileName string
	var metadataExists bool

	for _, entry := range entries {
		if strings.HasSuffix(entry.Name(), ".meta") {
			metadataPath := filepath.Join(metadataDir, entry.Name())
			data, err := os.ReadFile(metadataPath)
			if err != nil {
				continue
			}

			lines := strings.Split(string(data), "\n")
			var storedURL string

			for _, line := range lines {
				if strings.HasPrefix(line, "URL=") {
					storedURL = strings.TrimPrefix(line, "URL=")
				}
				if strings.HasPrefix(line, "ActualFileName=") {
					actualFileName = strings.TrimPrefix(line, "ActualFileName=")
				}
			}

			if storedURL == url {
				metadataExists = true
				break
			}
		}
	}

	if !metadataExists {
		return true, nil // 没有元数据，需要下载
	}

	// 如果没有实际文件名记录，回退到原来的扩展名检查
	if actualFileName == "" {
		expectedExt := ""
		if docType == "xlsx" {
			expectedExt = ".xlsx"
		} else if docType == "csv" {
			expectedExt = ".csv"
		} else {
			return true, nil
		}

		fileEntries, err := os.ReadDir(outputDir)
		if err != nil {
			return true, nil
		}

		for _, entry := range fileEntries {
			if !entry.IsDir() && strings.HasSuffix(strings.ToLower(entry.Name()), expectedExt) {
				return false, nil
			}
		}
		return true, nil
	}

	// 检查实际文件名是否存在
	actualFilePath := filepath.Join(outputDir, actualFileName)
	if _, err := os.Stat(actualFilePath); err == nil {
		return false, nil // 文件存在，不需要下载
	} else {
		return true, nil // 文件不存在，需要下载
	}
}

// 过滤需要同步的文档（用于增量模式）
func filterDocumentsForSync(ctx context.Context, client *core.Client, documents []DocConfig, outputDir string, syncSettings *SyncSettings) ([]DocConfig, error) {
	if syncSettings.SyncMode != "incremental" {
		return documents, nil
	}

	var needSync []DocConfig
	for _, doc := range documents {
		should, err := shouldSyncDocument(ctx, client, doc, outputDir, syncSettings)
		if err != nil {
			return nil, fmt.Errorf("检查文档 %s 同步状态失败: %v", doc.Name, err)
		}
		if should {
			needSync = append(needSync, doc)
		} else {
			fmt.Printf("跳过已存在文档: %s\n", doc.Name)
		}
	}
	return needSync, nil
}

// 保存带有实际文件名的文档元数据（用于表格文档的增量同步）
func saveDocumentMetadataWithFileName(ctx context.Context, client *core.Client, doc DocConfig, outputDir string, syncSettings *SyncSettings, actualFileName string) error {
	actualOutputDir := outputDir
	if syncSettings.OrganizeByGroup && doc.Group != "" {
		actualOutputDir = filepath.Join(outputDir, doc.Group)
	}

	// 创建元数据目录
	metadataDir := filepath.Join(actualOutputDir, ".feishu2md")

	// 确保元数据目录存在
	if err := os.MkdirAll(metadataDir, 0755); err != nil {
		return nil // 忽略错误
	}

	// 使用配置中的名称作为元数据文件名
	metadataFileName := fmt.Sprintf("%s.meta", utils.SanitizeFileName(doc.Name))
	metadataPath := filepath.Join(metadataDir, metadataFileName)

	// 保存元数据，只保留必要字段
	metadata := fmt.Sprintf("URL=%s\nName=%s\nActualFileName=%s\nSyncTime=%s\n",
		doc.URL, doc.Name, actualFileName, time.Now().Format(time.RFC3339))

	// 保存元数据文件
	err := os.WriteFile(metadataPath, []byte(metadata), 0644)
	if err != nil {
		fmt.Printf("Warning: failed to save metadata for %s: %v\n", doc.Name, err)
	}

	return nil
}

// 保存文档元数据（用于增量同步）
func saveDocumentMetadata(ctx context.Context, client *core.Client, doc DocConfig, outputDir string, syncSettings *SyncSettings) error {
	actualOutputDir := outputDir
	if syncSettings.OrganizeByGroup && doc.Group != "" {
		actualOutputDir = filepath.Join(outputDir, doc.Group)
	}

	// 获取文档信息
	docType, docToken, err := utils.ValidateDocumentURL(doc.URL)
	if err != nil {
		return nil // 忽略错误，不影响主要功能
	}

	// 创建元数据目录
	metadataDir := filepath.Join(actualOutputDir, ".feishu2md")

	// 确定元数据文件名
	var metadataFileName string
	var documentName string

	if syncSettings.UseOriginalTitle && docType == "docx" {
		// 使用原始标题的情况，需要获取文档标题
		docx, _, err := client.GetDocxContent(ctx, docToken)
		if err != nil {
			return nil // 忽略错误，不影响主要功能
		}
		documentName = docx.Title
		metadataFileName = fmt.Sprintf("%s.meta", utils.SanitizeFileName(docx.Title))
	} else {
		// 使用配置中的名称
		documentName = doc.Name
		metadataFileName = fmt.Sprintf("%s.meta", utils.SanitizeFileName(doc.Name))
	}

	metadataPath := filepath.Join(metadataDir, metadataFileName)

	if docType != "docx" {
	// 非docx文档，保存简化的元数据（暂时无法检测版本更新）
	// 对于非docx文档，实际文件名就是 documentName.md
	actualFileName := fmt.Sprintf("%s.md", documentName)
	metadata := fmt.Sprintf("URL=%s\nName=%s\nActualFileName=%s\nSyncTime=%s\n",
		doc.URL, doc.Name, actualFileName, time.Now().Format(time.RFC3339))		// 确保元数据目录存在
		if err := os.MkdirAll(metadataDir, 0755); err != nil {
			return nil // 忽略错误
		}

		// 保存元数据文件
		err = os.WriteFile(metadataPath, []byte(metadata), 0644)
		if err != nil {
			fmt.Printf("Warning: failed to save metadata for %s: %v\n", doc.Name, err)
		}
		return nil
	}

	// docx文档，保存包含RevisionID的元数据
	docx, _, err := client.GetDocxContent(ctx, docToken)
	if err != nil {
		return nil // 忽略错误，不影响主要功能
	}

	// 对于docx文档，实际文件名就是 documentName.md
	actualFileName := fmt.Sprintf("%s.md", documentName)
	
	// 创建元数据内容（包含RevisionID用于版本检测）
	metadata := fmt.Sprintf("URL=%s\nName=%s\nActualFileName=%s\nRevisionID=%d\nSyncTime=%s\n",
		doc.URL, doc.Name, actualFileName, docx.RevisionID, time.Now().Format(time.RFC3339))

	// 确保元数据目录存在
	if err := os.MkdirAll(metadataDir, 0755); err != nil {
		return nil // 忽略错误
	}

	// 保存元数据文件
	err = os.WriteFile(metadataPath, []byte(metadata), 0644)
	if err != nil {
		fmt.Printf("Warning: failed to save metadata for %s: %v\n", doc.Name, err)
	}

	return nil
}
