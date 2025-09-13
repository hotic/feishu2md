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
    Version   string        `json:"version" yaml:"version"`
    Sync      SyncSettings  `json:"sync" yaml:"sync"`
    Merge     MergeSettings `json:"merge" yaml:"merge"`
    Documents []DocConfig   `json:"documents" yaml:"documents"`
}

// SyncSettings represents sync-specific settings
type SyncSettings struct {
    OutputDir           string `json:"output_dir" yaml:"output_dir"`
    SyncMode            string `json:"sync_mode" yaml:"sync_mode"` // "clean_all" 或"incremental"
    ConcurrentDownloads int    `json:"concurrent_downloads" yaml:"concurrent_downloads"`
    OrganizeByGroup     bool   `json:"organize_by_group" yaml:"organize_by_group"`
    SkipImages          bool   `json:"skip_images" yaml:"skip_images"`
    UseOriginalTitle    bool   `json:"use_original_title" yaml:"use_original_title"`
    // 多维表格导出字段策略：
    // false（默认）导出表的全部字段；true 仅导出视图中“可见”的字段（更贴近飞书网页导出）
    BitableViewFieldsOnly bool `json:"bitable_view_fields_only" yaml:"bitable_view_fields_only"`
}

// MergeSettings represents merge-specific settings
type MergeSettings struct {
    InputDir             string   `json:"input_dir" yaml:"input_dir"`
    OutputDir            string   `json:"output_dir" yaml:"output_dir"`
    Filename             string   `json:"filename" yaml:"filename"`
    IncludeTimestamp     bool     `json:"include_timestamp" yaml:"include_timestamp"`
    SortFiles            bool     `json:"sort_files" yaml:"sort_files"`
    HeaderTitle          string   `json:"header_title" yaml:"header_title"`
    HeaderKeywords       []string `json:"header_keywords,omitempty" yaml:"header_keywords,omitempty"`
    GroupHeaderKeywords  []string `json:"group_header_keywords,omitempty" yaml:"group_header_keywords,omitempty"`
}

// DocConfig represents a single document configuration
// Type: optional doc type override:
//   - "docx" / "wiki" / "folder" keep existing behaviors
//   - "csv" / "xlsx" mean export Feishu Bitable as CSV/XLSX (requires table/view in URL)
type DocConfig struct {
    Name       string `json:"name" yaml:"name"`
    URL        string `json:"url" yaml:"url"`
    Group      string `json:"group,omitempty" yaml:"group,omitempty"`
    SkipImages *bool  `json:"skip_images,omitempty" yaml:"skip_images,omitempty"`
    Type       string `json:"type,omitempty" yaml:"type,omitempty"`
    // 针对单个文档覆盖：仅导出视图可见字段
    BitableViewFieldsOnly *bool `json:"bitable_view_fields_only,omitempty" yaml:"bitable_view_fields_only,omitempty"`
}

// NewSyncConfig creates a new sync configuration with defaults
func NewSyncConfig() *SyncConfig {
    return &SyncConfig{
        Version: "1.0",
        Sync: SyncSettings{
            OutputDir:           "./feishu_docs",
            SyncMode:            "clean_all",
            ConcurrentDownloads: 3,
            OrganizeByGroup:     true,
            SkipImages:          false,
            BitableViewFieldsOnly: false,
        },
        Merge: MergeSettings{
            InputDir:         "./feishu_docs",
            OutputDir:        "./",
            Filename:         "merged_docs.md",
            IncludeTimestamp: true,
            SortFiles:        true,
            HeaderTitle:      "合并的文档集",
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
    yamlPath := filepath.Join(configPath, "feishu2md", "config.yaml")
    if _, err := os.Stat(yamlPath); err == nil {
        return yamlPath, nil
    }
    ymlPath := filepath.Join(configPath, "feishu2md", "config.yml")
    if _, err := os.Stat(ymlPath); err == nil {
        return ymlPath, nil
    }
    jsonPath := filepath.Join(configPath, "feishu2md", "config.json")
    if _, err := os.Stat(jsonPath); err == nil {
        return jsonPath, nil
    }
    // Default to YAML for new configs
    return yamlPath, nil
}

// LoadSyncConfig loads sync configuration from file
func LoadSyncConfig(path string) (*SyncConfig, error) {
    if path == "" {
        // Preference order: local config.yml/config.yaml, then legacy sync_config files, then user config dir
        if _, err := os.Stat("config.yml"); err == nil {
            path = "config.yml"
        } else if _, err := os.Stat("config.yaml"); err == nil {
            path = "config.yaml"
        } else if _, err := os.Stat("sync_config.yaml"); err == nil {
            path = "sync_config.yaml"
        } else if _, err := os.Stat("sync_config.yml"); err == nil {
            path = "sync_config.yml"
        } else {
            var err error
            path, err = GetSyncConfigPath()
            if err != nil {
                return nil, err
            }
        }
    }

    // If path doesn't have extension, try both YAML and JSON
    if !strings.HasSuffix(path, ".yaml") && !strings.HasSuffix(path, ".yml") && !strings.HasSuffix(path, ".json") {
        yamlPath := path + ".yaml"
        if _, err := os.Stat(yamlPath); err == nil {
            path = yamlPath
        } else {
            ymlPath := path + ".yml"
            if _, err := os.Stat(ymlPath); err == nil {
                path = ymlPath
            } else {
                jsonPath := path + ".json"
                if _, err := os.Stat(jsonPath); err == nil {
                    path = jsonPath
                } else {
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

    // Try by name
    for i, doc := range c.Documents {
        if doc.Name == nameOrIndex {
            c.Documents = append(c.Documents[:i], c.Documents[i+1:]...)
            return nil
        }
    }

    return fmt.Errorf("document %s not found", nameOrIndex)
}

// GetDocuments returns all documents, optionally filtered by group
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
