package paths

import (
	"os"
	"path/filepath"
)

const (
	DefaultOutputDir     = `C:\Ziitech\Menu`
	DefaultTemplateDir   = `C:\Ziitech`
	DefaultTemplateFile  = `C:\Ziitech\ZiiPOS_MenuTemplate.xlsx`
	TemplateDownloadURL  = "https://download.ziicloud.com/other/ZiiPOS_MenuTemplate.xlsx"
)

func ConfigDir() string {
	if appData := os.Getenv("APPDATA"); appData != "" {
		return filepath.Join(appData, "ZiiPOSMenuConverter")
	}
	home, _ := os.UserHomeDir()
	return filepath.Join(home, "ZiiPOSMenuConverter")
}

func ConfigFile() string {
	return filepath.Join(ConfigDir(), "settings_v4_go.json")
}
