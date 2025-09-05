package core

import (
	"bytes"
	"context"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"time"

	"github.com/chyroc/lark"
	"github.com/chyroc/lark_rate_limiter"
)

type Client struct {
	larkClient *lark.Lark
}

func NewClient(appID, appSecret string) *Client {
	return &Client{
		larkClient: lark.New(
			lark.WithAppCredential(appID, appSecret),
			lark.WithTimeout(60*time.Second),
			lark.WithApiMiddleware(lark_rate_limiter.Wait(4, 4)),
		),
	}
}

func (c *Client) DownloadImage(ctx context.Context, imgToken, outDir string) (string, error) {
	resp, _, err := c.larkClient.Drive.DownloadDriveMedia(ctx, &lark.DownloadDriveMediaReq{
		FileToken: imgToken,
	})
	if err != nil {
		return imgToken, err
	}
	fileext := filepath.Ext(resp.Filename)
	filename := fmt.Sprintf("%s/%s%s", outDir, imgToken, fileext)
	err = os.MkdirAll(filepath.Dir(filename), 0o755)
	if err != nil {
		return imgToken, err
	}
	file, err := os.OpenFile(filename, os.O_CREATE|os.O_WRONLY, 0o666)
	if err != nil {
		return imgToken, err
	}
	defer file.Close()
	_, err = io.Copy(file, resp.File)
	if err != nil {
		return imgToken, err
	}
	return filename, nil
}

func (c *Client) DownloadImageRaw(ctx context.Context, imgToken, imgDir string) (string, []byte, error) {
	resp, _, err := c.larkClient.Drive.DownloadDriveMedia(ctx, &lark.DownloadDriveMediaReq{
		FileToken: imgToken,
	})
	if err != nil {
		return imgToken, nil, err
	}
	fileext := filepath.Ext(resp.Filename)
	filename := fmt.Sprintf("%s/%s%s", imgDir, imgToken, fileext)
	buf := new(bytes.Buffer)
	buf.ReadFrom(resp.File)
	return filename, buf.Bytes(), nil
}

func (c *Client) GetDocxContent(ctx context.Context, docToken string) (*lark.DocxDocument, []*lark.DocxBlock, error) {
	resp, _, err := c.larkClient.Drive.GetDocxDocument(ctx, &lark.GetDocxDocumentReq{
		DocumentID: docToken,
	})
	if err != nil {
		return nil, nil, err
	}
	docx := &lark.DocxDocument{
		DocumentID: resp.Document.DocumentID,
		RevisionID: resp.Document.RevisionID,
		Title:      resp.Document.Title,
	}
	var blocks []*lark.DocxBlock
	var pageToken *string
	for {
		resp2, _, err := c.larkClient.Drive.GetDocxBlockListOfDocument(ctx, &lark.GetDocxBlockListOfDocumentReq{
			DocumentID: docx.DocumentID,
			PageToken:  pageToken,
		})
		if err != nil {
			return docx, nil, err
		}
		blocks = append(blocks, resp2.Items...)
		pageToken = &resp2.PageToken
		if !resp2.HasMore {
			break
		}
	}
	return docx, blocks, nil
}

func (c *Client) GetWikiNodeInfo(ctx context.Context, token string) (*lark.GetWikiNodeRespNode, error) {
	resp, _, err := c.larkClient.Drive.GetWikiNode(ctx, &lark.GetWikiNodeReq{
		Token: token,
	})
	if err != nil {
		return nil, err
	}
	return resp.Node, nil
}

func (c *Client) GetDriveFolderFileList(ctx context.Context, pageToken *string, folderToken *string) ([]*lark.GetDriveFileListRespFile, error) {
	resp, _, err := c.larkClient.Drive.GetDriveFileList(ctx, &lark.GetDriveFileListReq{
		PageSize:    nil,
		PageToken:   pageToken,
		FolderToken: folderToken,
	})
	if err != nil {
		return nil, err
	}
	files := resp.Files
	for resp.HasMore {
		resp, _, err = c.larkClient.Drive.GetDriveFileList(ctx, &lark.GetDriveFileListReq{
			PageSize:    nil,
			PageToken:   &resp.NextPageToken,
			FolderToken: folderToken,
		})
		if err != nil {
			return nil, err
		}
		files = append(files, resp.Files...)
	}
	return files, nil
}

func (c *Client) GetWikiName(ctx context.Context, spaceID string) (string, error) {
	resp, _, err := c.larkClient.Drive.GetWikiSpace(ctx, &lark.GetWikiSpaceReq{
		SpaceID: spaceID,
	})

	if err != nil {
		return "", err
	}

	return resp.Space.Name, nil
}

func (c *Client) GetWikiNodeList(ctx context.Context, spaceID string, parentNodeToken *string) ([]*lark.GetWikiNodeListRespItem, error) {
	resp, _, err := c.larkClient.Drive.GetWikiNodeList(ctx, &lark.GetWikiNodeListReq{
		SpaceID:         spaceID,
		PageSize:        nil,
		PageToken:       nil,
		ParentNodeToken: parentNodeToken,
	})

	if err != nil {
		return nil, err
	}

	nodes := resp.Items
	previousPageToken := ""

	for resp.HasMore && previousPageToken != resp.PageToken {
		previousPageToken = resp.PageToken
		resp, _, err := c.larkClient.Drive.GetWikiNodeList(ctx, &lark.GetWikiNodeListReq{
			SpaceID:         spaceID,
			PageSize:        nil,
			PageToken:       &resp.PageToken,
			ParentNodeToken: parentNodeToken,
		})

		if err != nil {
			return nil, err
		}

		nodes = append(nodes, resp.Items...)
	}

	return nodes, nil
}

func (c *Client) GetBitableMeta(ctx context.Context, appToken string) (*lark.GetBitableMetaRespApp, error) {
	resp, _, err := c.larkClient.Bitable.GetBitableMeta(ctx, &lark.GetBitableMetaReq{
		AppToken: appToken,
	})
	if err != nil {
		return nil, err
	}
	return resp.App, nil
}

func (c *Client) GetBitableTableList(ctx context.Context, appToken string) ([]*lark.GetBitableTableListRespItem, error) {
	var all []*lark.GetBitableTableListRespItem
	var pageToken *string
	for {
		resp, _, err := c.larkClient.Bitable.GetBitableTableList(ctx, &lark.GetBitableTableListReq{
			AppToken:  appToken,
			PageToken: pageToken,
			PageSize:  nil,
		})
		if err != nil {
			return nil, err
		}
		all = append(all, resp.Items...)
		if !resp.HasMore || resp.PageToken == "" || (pageToken != nil && *pageToken == resp.PageToken) {
			break
		}
		pageToken = &resp.PageToken
	}
	return all, nil
}

func (c *Client) GetBitableViewList(ctx context.Context, appToken, tableID string) ([]*lark.GetBitableViewListRespItem, error) {
	var all []*lark.GetBitableViewListRespItem
	var pageToken *string
	for {
		resp, _, err := c.larkClient.Bitable.GetBitableViewList(ctx, &lark.GetBitableViewListReq{
			AppToken:  appToken,
			TableID:   tableID,
			PageSize:  nil,
			PageToken: pageToken,
		})
		if err != nil {
			return nil, err
		}
		all = append(all, resp.Items...)
		if !resp.HasMore || resp.PageToken == "" || (pageToken != nil && *pageToken == resp.PageToken) {
			break
		}
		pageToken = &resp.PageToken
	}
	return all, nil
}

func (c *Client) GetBitableFieldList(ctx context.Context, appToken, tableID string, viewID *string) ([]*lark.GetBitableFieldListRespItem, error) {
	var all []*lark.GetBitableFieldListRespItem
	var pageToken *string
	for {
		req := &lark.GetBitableFieldListReq{
			AppToken:  appToken,
			TableID:   tableID,
			ViewID:    viewID,
			PageToken: pageToken,
		}
		resp, _, err := c.larkClient.Bitable.GetBitableFieldList(ctx, req)
		if err != nil {
			return nil, err
		}
		all = append(all, resp.Items...)
		if !resp.HasMore || resp.PageToken == "" || (pageToken != nil && *pageToken == resp.PageToken) {
			break
		}
		pageToken = &resp.PageToken
	}
	return all, nil
}

func (c *Client) GetBitableRecordPage(ctx context.Context, appToken, tableID string, viewID *string, pageToken *string, pageSize int64) (*lark.GetBitableRecordListResp, error) {
	req := &lark.GetBitableRecordListReq{
		AppToken:  appToken,
		TableID:   tableID,
		ViewID:    viewID,
		PageToken: pageToken,
		PageSize:  &pageSize,
	}
	return c.getBitableRecordList(ctx, req)
}

func (c *Client) getBitableRecordList(ctx context.Context, req *lark.GetBitableRecordListReq) (*lark.GetBitableRecordListResp, error) {
	resp, _, err := c.larkClient.Bitable.GetBitableRecordList(ctx, req)
	if err != nil {
		return nil, err
	}
	return resp, nil
}
