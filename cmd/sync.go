package main

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/Wsine/feishu2md/core"
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
	fmt.Printf("Clean Before Sync: %v\n", config.Sync.CleanBeforeSync)

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

	// Clean output directory if requested
	// CleanBeforeSync: always clean when true
	// --force flag: force clean even if CleanBeforeSync is false
	if syncConfig.Sync.CleanBeforeSync || syncOpts.force {
		fmt.Println("Cleaning output directory...")
		if err := cleanOutputDirectory(syncConfig.Sync.OutputDir); err != nil {
			fmt.Printf("Warning: failed to clean output directory: %v\n", err)
		}
	}

	// Create client
	client := core.NewClient(feishuConfig.Feishu.AppId, feishuConfig.Feishu.AppSecret)
	ctx2 := context.Background()

	// Sync documents with concurrency control
	var wg sync.WaitGroup
	semaphore := make(chan struct{}, syncConfig.Sync.ConcurrentDownloads)
	errors := make([]error, 0)
	var errorsMux sync.Mutex

	startTime := time.Now()
	successCount := 0
	var successMux sync.Mutex

	for _, doc := range documents {
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
	fmt.Printf("成功: %d/%d\n", successCount, len(documents))

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

	opts := DownloadOpts{
		outputDir:  outputDir,
		dump:       false,
		batch:      docType == "folder",
		wiki:       docType == "wiki_space",
		docName:    doc.Name, // Pass the document name from config
		skipImages: skipImages,
	}

	switch docType {
	case "wiki_space":
		// For entire wiki spaces (with /wiki/settings/ URL)
		return downloadWiki(ctx, client, doc.URL)
	case "wiki_page":
		// For individual wiki pages, treat them as documents
		return downloadDocument(ctx, client, doc.URL, &opts)
	case "folder":
		return downloadDocuments(ctx, client, doc.URL)
	case "csv":
		return exportBitable(ctx, client, doc.URL, "csv", outputDir, doc.Name)
	case "xlsx":
		return exportBitable(ctx, client, doc.URL, "xlsx", outputDir, doc.Name)
	default: // docx
		return downloadDocument(ctx, client, doc.URL, &opts)
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
