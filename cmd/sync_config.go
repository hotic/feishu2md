package main

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"gopkg.in/yaml.v3"
)

// SyncConfig represents the sync configuration structure
type SyncConfig struct {
	Version   string       `json:"version" yaml:"version"`
	Sync      SyncSettings `json:"sync" yaml:"sync"`
	Documents []DocConfig  `json:"documents" yaml:"documents"`
}

// SyncSettings represents sync-specific settings
type SyncSettings struct {
	OutputDir           string `json:"output_dir" yaml:"output_dir"`                     // 输出目录
	CleanBeforeSync     bool   `json:"clean_before_sync" yaml:"clean_before_sync"`       // 同步前是否清空目录
	ConcurrentDownloads int    `json:"concurrent_downloads" yaml:"concurrent_downloads"` // 并发下载数
	OrganizeByGroup     bool   `json:"organize_by_group" yaml:"organize_by_group"`       // 是否按组织结构存储
	SkipImages          bool   `json:"skip_images" yaml:"skip_images"`                   // 是否跳过图片下载（全局配置）
}

// DocConfig represents a single document configuration
type DocConfig struct {
	Name       string `json:"name" yaml:"name"`                                   // 文档名称
	URL        string `json:"url" yaml:"url"`                                     // 文档URL
	Group      string `json:"group,omitempty" yaml:"group,omitempty"`             // 文档分组（可选）
	SkipImages *bool  `json:"skip_images,omitempty" yaml:"skip_images,omitempty"` // 是否跳过图片下载（单文档配置，使用指针以区分是否设置）
}

// NewSyncConfig creates a new sync configuration with defaults
func NewSyncConfig() *SyncConfig {
	return &SyncConfig{
		Version: "1.0",
		Sync: SyncSettings{
			OutputDir:           "./feishu_docs",
			CleanBeforeSync:     false,
			ConcurrentDownloads: 3,
			OrganizeByGroup:     true,
			SkipImages:          false, // 默认不跳过图片下载
		},
		Documents: []DocConfig{},
	}
}

// GetSyncConfigPath returns the default sync config file path
func GetSyncConfigPath() (string, error) {
	configPath, err := os.UserConfigDir()
	if err != nil {
		return "", err
	}
	// Check for YAML file first, then JSON
	yamlPath := filepath.Join(configPath, "feishu2md", "sync_config.yaml")
	if _, err := os.Stat(yamlPath); err == nil {
		return yamlPath, nil
	}
	jsonPath := filepath.Join(configPath, "feishu2md", "sync_config.json")
	if _, err := os.Stat(jsonPath); err == nil {
		return jsonPath, nil
	}
	// Default to YAML for new configs
	return yamlPath, nil
}

// LoadSyncConfig loads sync configuration from file
func LoadSyncConfig(path string) (*SyncConfig, error) {
	if path == "" {
		// 优先查找当前目录的配置文件
		// 1. 尝试当前目录的 sync_config.yaml
		if _, err := os.Stat("sync_config.yaml"); err == nil {
			path = "sync_config.yaml"
		} else if _, err := os.Stat("sync_config.yml"); err == nil {
			// 2. 尝试当前目录的 sync_config.yml
			path = "sync_config.yml"
		} else {
			// 3. 使用用户配置目录
			var err error
			path, err = GetSyncConfigPath()
			if err != nil {
				return nil, err
			}
		}
	}

	// If path doesn't have extension, try both YAML and JSON
	if !strings.HasSuffix(path, ".yaml") && !strings.HasSuffix(path, ".yml") && !strings.HasSuffix(path, ".json") {
		// Try YAML first
		yamlPath := path + ".yaml"
		if _, err := os.Stat(yamlPath); err == nil {
			path = yamlPath
		} else {
			// Try YML
			ymlPath := path + ".yml"
			if _, err := os.Stat(ymlPath); err == nil {
				path = ymlPath
			} else {
				// Try JSON
				jsonPath := path + ".json"
				if _, err := os.Stat(jsonPath); err == nil {
					path = jsonPath
				} else {
					// Default to YAML
					path = yamlPath
				}
			}
		}
	}

	data, err := os.ReadFile(path)
	if err != nil {
		if os.IsNotExist(err) {
			// Return new config if file doesn't exist
			return NewSyncConfig(), nil
		}
		return nil, err
	}

	var config SyncConfig

	// Determine format by extension or content
	if strings.HasSuffix(path, ".yaml") || strings.HasSuffix(path, ".yml") {
		if err := yaml.Unmarshal(data, &config); err != nil {
			return nil, fmt.Errorf("invalid YAML format: %v", err)
		}
	} else if strings.HasSuffix(path, ".json") {
		if err := json.Unmarshal(data, &config); err != nil {
			return nil, fmt.Errorf("invalid JSON format: %v", err)
		}
	} else {
		// Try to auto-detect format
		if err := yaml.Unmarshal(data, &config); err == nil {
			// Successfully parsed as YAML
		} else if err := json.Unmarshal(data, &config); err == nil {
			// Successfully parsed as JSON
		} else {
			return nil, fmt.Errorf("unable to parse config file as YAML or JSON")
		}
	}

	// Set defaults for missing values
	if config.Sync.ConcurrentDownloads <= 0 {
		config.Sync.ConcurrentDownloads = 3
	}
	if config.Sync.OutputDir == "" {
		config.Sync.OutputDir = "./feishu_docs"
	}

	return &config, nil
}

// SaveSyncConfig saves sync configuration to file
func (c *SyncConfig) Save(path string) error {
	if path == "" {
		var err error
		path, err = GetSyncConfigPath()
		if err != nil {
			return err
		}
	}

	// If path doesn't have extension, default to YAML
	if !strings.HasSuffix(path, ".yaml") && !strings.HasSuffix(path, ".yml") && !strings.HasSuffix(path, ".json") {
		path = path + ".yaml"
	}

	// Ensure directory exists
	dir := filepath.Dir(path)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return err
	}

	var data []byte
	var err error

	// Save based on extension
	if strings.HasSuffix(path, ".yaml") || strings.HasSuffix(path, ".yml") {
		data, err = yaml.Marshal(c)
		if err != nil {
			return err
		}
	} else if strings.HasSuffix(path, ".json") {
		data, err = json.MarshalIndent(c, "", "  ")
		if err != nil {
			return err
		}
	} else {
		// Default to YAML
		data, err = yaml.Marshal(c)
		if err != nil {
			return err
		}
	}

	return os.WriteFile(path, data, 0644)
}

// AddDocument adds a new document to the configuration
func (c *SyncConfig) AddDocument(name, url, group string) error {
	// Check for duplicates
	for _, doc := range c.Documents {
		if doc.URL == url {
			return fmt.Errorf("document with URL %s already exists", url)
		}
		if doc.Name == name {
			return fmt.Errorf("document with name %s already exists", name)
		}
	}

	c.Documents = append(c.Documents, DocConfig{
		Name:  name,
		URL:   url,
		Group: group,
	})

	return nil
}

// RemoveDocument removes a document by name or index
func (c *SyncConfig) RemoveDocument(nameOrIndex string) error {
	// Try to parse as index first
	var index int
	if _, err := fmt.Sscanf(nameOrIndex, "%d", &index); err == nil {
		if index >= 0 && index < len(c.Documents) {
			c.Documents = append(c.Documents[:index], c.Documents[index+1:]...)
			return nil
		}
		return fmt.Errorf("index %d out of range", index)
	}

	// Try to match by name
	for i, doc := range c.Documents {
		if doc.Name == nameOrIndex {
			c.Documents = append(c.Documents[:i], c.Documents[i+1:]...)
			return nil
		}
	}

	return fmt.Errorf("document %s not found", nameOrIndex)
}

// GetDocuments returns all documents, optionally filtered by group
// Documents commented out in YAML won't be included
func (c *SyncConfig) GetDocuments(group string) []DocConfig {
	var docs []DocConfig
	for _, doc := range c.Documents {
		if group != "" && doc.Group != group {
			continue
		}
		docs = append(docs, doc)
	}
	return docs
}

func contains(s, substr string) bool {
	return len(s) > 0 && len(substr) > 0 && len(s) >= len(substr) &&
		(s == substr || (len(s) > len(substr) &&
			(s[0:len(substr)] == substr ||
				s[len(s)-len(substr):] == substr ||
				(len(s) > len(substr) && containsInMiddle(s, substr)))))
}

func containsInMiddle(s, substr string) bool {
	if len(s) <= len(substr) {
		return false
	}
	for i := 1; i < len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return true
		}
	}
	return false
}
