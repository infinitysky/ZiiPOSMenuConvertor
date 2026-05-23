package template

import (
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"

	"menu-converter-v4-go/internal/paths"
)

func Ensure(path string) error {
	if _, err := os.Stat(path); err == nil {
		return nil
	}
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return err
	}
	resp, err := http.Get(paths.TemplateDownloadURL)
	if err != nil {
		return fmt.Errorf("%s: %w", paths.TemplateDownloadURL, err)
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("download failed: HTTP %d", resp.StatusCode)
	}
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	if _, err := io.Copy(f, resp.Body); err != nil {
		return err
	}
	return nil
}

func DefaultPath() string {
	return paths.DefaultTemplateFile
}
