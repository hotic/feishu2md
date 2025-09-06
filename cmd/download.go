package main

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"

	"github.com/88250/lute"
	"github.com/Wsine/feishu2md/core"
	"github.com/Wsine/feishu2md/utils"
	"github.com/chyroc/lark"
	"github.com/pkg/errors"
)

type DownloadOpts struct {
	outputDir        string
	dump             bool
	batch            bool
	wiki             bool
	docName          string // Optional custom document name
	skipImages       bool   // 是否跳过图片下载
	useOriginalTitle bool   // Whether to use original title instead of docName
}

var dlOpts = DownloadOpts{}
var dlConfig core.Config

func downloadDocument(ctx context.Context, client *core.Client, url string, opts *DownloadOpts) (string, error) {
	// Validate the url to download
	docType, docToken, err := utils.ValidateDocumentURL(url)
	if err != nil {
		return "", err
	}
	fmt.Println("获取文档令牌:", docToken)

	// for a wiki page, we need to renew docType and docToken first
	if docType == "wiki" {
		node, err := client.GetWikiNodeInfo(ctx, docToken)
		if err != nil {
			err = fmt.Errorf("GetWikiNodeInfo err: %v for %v", err, url)
		}
		utils.CheckErr(err)
		docType = node.ObjType
		docToken = node.ObjToken
	}
	if docType == "docs" {
		return "", errors.Errorf(
			`Feishu Docs is no longer supported. ` +
				`Please refer to the Readme/Release for v1_support.`)
	}

	// Process the download
	docx, blocks, err := client.GetDocxContent(ctx, docToken)
	utils.CheckErr(err)

	parser := core.NewParser(dlConfig.Output)

	// Collect @mention user OpenIDs, resolve to display names, and set on parser
	collectMentionOpenIDs := func(blocks []*lark.DocxBlock) []string {
		ids := make([]string, 0)
		seen := make(map[string]struct{})
		addFromText := func(t *lark.DocxBlockText) {
			if t == nil {
				return
			}
			for _, el := range t.Elements {
				if el != nil && el.MentionUser != nil {
					id := el.MentionUser.UserID
					if id != "" {
						if _, ok := seen[id]; !ok {
							seen[id] = struct{}{}
							ids = append(ids, id)
						}
					}
				}
			}
		}
		for _, b := range blocks {
			if b.Page != nil {
				addFromText(b.Page)
			}
			if b.Text != nil {
				addFromText(b.Text)
			}
			if b.Heading1 != nil {
				addFromText(b.Heading1)
			}
			if b.Heading2 != nil {
				addFromText(b.Heading2)
			}
			if b.Heading3 != nil {
				addFromText(b.Heading3)
			}
			if b.Heading4 != nil {
				addFromText(b.Heading4)
			}
			if b.Heading5 != nil {
				addFromText(b.Heading5)
			}
			if b.Heading6 != nil {
				addFromText(b.Heading6)
			}
			if b.Heading7 != nil {
				addFromText(b.Heading7)
			}
			if b.Heading8 != nil {
				addFromText(b.Heading8)
			}
			if b.Heading9 != nil {
				addFromText(b.Heading9)
			}
			if b.Bullet != nil {
				addFromText(b.Bullet)
			}
			if b.Ordered != nil {
				addFromText(b.Ordered)
			}
			if b.Code != nil {
				addFromText(b.Code)
			}
			if b.Quote != nil {
				addFromText(b.Quote)
			}
			if b.Equation != nil {
				addFromText(b.Equation)
			}
			if b.Todo != nil {
				addFromText(b.Todo)
			}
		}
		return ids
	}
	mentionIDs := collectMentionOpenIDs(blocks)
	if len(mentionIDs) > 0 {
		fmt.Printf("  发现 %d 个 @提及用户，开始解析...\n", len(mentionIDs))
		nameMap := client.ResolveUserNames(ctx, mentionIDs)
		parser.SetMentionUserMap(nameMap)
		// Debug summary to help diagnose permission/config issues
		resolved := 0
		unresolvedList := make([]string, 0)
		for _, id := range mentionIDs {
			if nameMap[id] != "" {
				resolved++
			} else {
				unresolvedList = append(unresolvedList, id)
			}
		}
		if len(unresolvedList) > 0 {
			fmt.Printf("  @提及解析: %d/%d 成功，未解析: %v\n", resolved, len(mentionIDs), unresolvedList)
			if resolved == 0 {
				fmt.Printf("  💡 提示: 要获取正确的用户名，请:\n")
				fmt.Printf("     1. 在飞书开放平台为应用添加 'contact:user.base:readonly' 权限\n")
				fmt.Printf("     2. 或使用飞书网页版导出文档功能 (文件 > 导出 > Word)\n")
			}
		} else {
			fmt.Printf("  @提及解析: %d/%d 成功\n", resolved, len(mentionIDs))
		}
	}

	title := docx.Title
	markdown := parser.ParseDocxContent(docx, blocks)

	// Determine document name for image folder
	var docName string
	if opts.useOriginalTitle {
		// 使用飞书文档的原始标题
		docName = utils.SanitizeFileName(title)
	} else if opts.docName != "" {
		// Use the provided document name from config
		docName = utils.SanitizeFileName(opts.docName)
	} else if dlConfig.Output.TitleAsFilename {
		// Use title as folder name if configured
		docName = utils.SanitizeFileName(title)
	} else {
		// Default to token as folder name
		docName = docToken
	}

	// 检查是否跳过图片下载：opts.skipImages 优先于配置文件中的设置
	shouldSkipImages := opts.skipImages || dlConfig.Output.SkipImgDownload

	if !shouldSkipImages {
		// Create document-specific image directory
		imageDir := filepath.Join(opts.outputDir, docName)

		for _, imgToken := range parser.ImgTokens {
			localLink, err := client.DownloadImage(
				ctx, imgToken, imageDir,
			)
			if err != nil {
				return "", err
			}
			// Update the image path to be relative to the markdown file
			relPath := filepath.Join(docName, filepath.Base(localLink))
			markdown = strings.Replace(markdown, imgToken, relPath, 1)
		}
	} else {
		fmt.Printf("  跳过图片下载（共 %d 张图片）\n", len(parser.ImgTokens))
	}

	// Format the markdown document
	engine := lute.New(func(l *lute.Lute) {
		l.RenderOptions.AutoSpace = true
	})
	result := engine.FormatStr("md", markdown)

	// Handle the output directory and name
	if _, err := os.Stat(opts.outputDir); os.IsNotExist(err) {
		if err := os.MkdirAll(opts.outputDir, 0o755); err != nil {
			return "", err
		}
	}

	if dlOpts.dump {
		jsonName := fmt.Sprintf("%s.json", docToken)
		outputPath := filepath.Join(opts.outputDir, jsonName)
		data := struct {
			Document *lark.DocxDocument `json:"document"`
			Blocks   []*lark.DocxBlock  `json:"blocks"`
		}{
			Document: docx,
			Blocks:   blocks,
		}
		pdata := utils.PrettyPrint(data)

		if err = os.WriteFile(outputPath, []byte(pdata), 0o644); err != nil {
			return "", err
		}
		fmt.Printf("Dumped json response to %s\n", outputPath)
	}

	// Write to markdown file
	var mdName string
	if opts.useOriginalTitle {
		// 使用飞书文档的原始标题
		mdName = fmt.Sprintf("%s.md", utils.SanitizeFileName(title))
	} else if opts.docName != "" {
		// Use the provided document name from config
		mdName = fmt.Sprintf("%s.md", utils.SanitizeFileName(opts.docName))
	} else if dlConfig.Output.TitleAsFilename {
		// Use title as filename if configured
		mdName = fmt.Sprintf("%s.md", utils.SanitizeFileName(title))
	} else {
		// Default to token as filename
		mdName = fmt.Sprintf("%s.md", docToken)
	}
	outputPath := filepath.Join(opts.outputDir, mdName)
	if err = os.WriteFile(outputPath, []byte(result), 0o644); err != nil {
		return "", err
	}
	fmt.Printf("已下载 markdown 文件到 %s\n", outputPath)

	return mdName, nil
}

func downloadDocuments(ctx context.Context, client *core.Client, url string) error {
	// Validate the url to download
	folderToken, err := utils.ValidateFolderURL(url)
	if err != nil {
		return err
	}
	fmt.Println("Captured folder token:", folderToken)

	// Error channel and wait group
	errChan := make(chan error)
	wg := sync.WaitGroup{}

	// Recursively go through the folder and download the documents
	var processFolder func(ctx context.Context, folderPath, folderToken string) error
	processFolder = func(ctx context.Context, folderPath, folderToken string) error {
		files, err := client.GetDriveFolderFileList(ctx, nil, &folderToken)
		if err != nil {
			return err
		}
		for _, file := range files {
			if file.Type == "folder" {
				_folderPath := filepath.Join(folderPath, file.Name)
				if err := processFolder(ctx, _folderPath, file.Token); err != nil {
					return err
				}
			} else if file.Type == "docx" {
				// Use file name as document name for image folder
				opts := DownloadOpts{
					outputDir:        folderPath,
					dump:             dlOpts.dump,
					batch:            false,
					docName:          file.Name,
					skipImages:       dlOpts.skipImages, // 继承父级的skipImages设置
					useOriginalTitle: false,             // 在folder下载中使用文件名，不使用原始标题
				}
				// concurrently download the document
				wg.Add(1)
				go func(_url string) {
					if _, err := downloadDocument(ctx, client, _url, &opts); err != nil {
						errChan <- err
					}
					wg.Done()
				}(file.URL)
			}
		}
		return nil
	}
	if err := processFolder(ctx, dlOpts.outputDir, folderToken); err != nil {
		return err
	}

	// Wait for all the downloads to finish
	go func() {
		wg.Wait()
		close(errChan)
	}()
	for err := range errChan {
		return err
	}
	return nil
}

func downloadWiki(ctx context.Context, client *core.Client, url string) error {
	prefixURL, spaceID, err := utils.ValidateWikiURL(url)
	if err != nil {
		return err
	}

	folderPath, err := client.GetWikiName(ctx, spaceID)
	if err != nil {
		return err
	}
	if folderPath == "" {
		return fmt.Errorf("failed to GetWikiName")
	}

	errChan := make(chan error)

	var maxConcurrency = 10 // Set the maximum concurrency level
	wg := sync.WaitGroup{}
	semaphore := make(chan struct{}, maxConcurrency) // Create a semaphore with the maximum concurrency level

	var downloadWikiNode func(ctx context.Context,
		client *core.Client,
		spaceID string,
		parentPath string,
		parentNodeToken *string) error

	downloadWikiNode = func(ctx context.Context,
		client *core.Client,
		spaceID string,
		folderPath string,
		parentNodeToken *string) error {
		nodes, err := client.GetWikiNodeList(ctx, spaceID, parentNodeToken)
		if err != nil {
			return err
		}
		for _, n := range nodes {
			if n.HasChild {
				_folderPath := filepath.Join(folderPath, n.Title)
				if err := downloadWikiNode(ctx, client,
					spaceID, _folderPath, &n.NodeToken); err != nil {
					return err
				}
			}
			if n.ObjType == "docx" {
				// Use node title as document name for image folder
				opts := DownloadOpts{
					outputDir:        folderPath,
					dump:             dlOpts.dump,
					batch:            false,
					docName:          n.Title,
					skipImages:       dlOpts.skipImages, // 继承父级的skipImages设置
					useOriginalTitle: false,             // 在wiki下载中使用节点标题，不使用原始标题
				}
				wg.Add(1)
				semaphore <- struct{}{}
				go func(_url string) {
					if _, err := downloadDocument(ctx, client, _url, &opts); err != nil {
						errChan <- err
					}
					wg.Done()
					<-semaphore
				}(prefixURL + "/wiki/" + n.NodeToken)
				// downloadDocument(ctx, client, prefixURL+"/wiki/"+n.NodeToken, &opts)
			}
		}
		return nil
	}

	if err = downloadWikiNode(ctx, client, spaceID, folderPath, nil); err != nil {
		return err
	}

	// Wait for all the downloads to finish
	go func() {
		wg.Wait()
		close(errChan)
	}()
	for err := range errChan {
		return err
	}
	return nil
}

func handleDownloadCommand(url string) error {
	// Load config
	configPath, err := core.GetConfigFilePath()
	if err != nil {
		return err
	}
	config, err := core.ReadConfigFromFile(configPath)
	if err != nil {
		return err
	}
	dlConfig = *config

	// Instantiate the client
	client := core.NewClient(
		dlConfig.Feishu.AppId, dlConfig.Feishu.AppSecret,
	)
	ctx := context.Background()

	if dlOpts.batch {
		return downloadDocuments(ctx, client, url)
	}

	if dlOpts.wiki {
		return downloadWiki(ctx, client, url)
	}

	_, err = downloadDocument(ctx, client, url, &dlOpts)
	return err
}
