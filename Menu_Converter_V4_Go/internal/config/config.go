package config

import (
	"encoding/json"
	"os"

	"menu-converter-v4-go/internal/paths"
)

type Settings struct {
	UILang     string `json:"ui_lang"`
	OutputPath string `json:"output_path,omitempty"`
}

func Load() Settings {
	cfg := Settings{UILang: "en", OutputPath: paths.DefaultOutputDir}
	data, err := os.ReadFile(paths.ConfigFile())
	if err != nil {
		return cfg
	}
	_ = json.Unmarshal(data, &cfg)
	if cfg.UILang == "" {
		cfg.UILang = "en"
	}
	if cfg.OutputPath == "" {
		cfg.OutputPath = paths.DefaultOutputDir
	}
	return cfg
}

func Save(cfg Settings) error {
	if err := os.MkdirAll(paths.ConfigDir(), 0o755); err != nil {
		return err
	}
	data, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(paths.ConfigFile(), data, 0o644)
}
