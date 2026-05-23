package server

import (
	"encoding/json"
	"io/fs"
	"net/http"
	"os"
	"path/filepath"
	"strings"

	"github.com/sqweek/dialog"

	"menu-converter-v4-go/internal/config"
	"menu-converter-v4-go/internal/convert"
	"menu-converter-v4-go/internal/i18n"
	"menu-converter-v4-go/internal/pe"
	"menu-converter-v4-go/internal/template"
)

type Server struct {
	cfg     config.Settings
	webFS   fs.FS
	peItems []pe.Item
}

func New(webFS fs.FS, cfg config.Settings) *Server {
	return &Server{webFS: webFS, cfg: cfg}
}

func (s *Server) Handler() http.Handler {
	mux := http.NewServeMux()
	mux.Handle("/", http.FileServer(http.FS(s.webFS)))
	mux.HandleFunc("/api/config", s.handleConfig)
	mux.HandleFunc("/api/i18n/", s.handleI18n)
	mux.HandleFunc("/api/dialog/file", s.handleDialogFile)
	mux.HandleFunc("/api/dialog/dir", s.handleDialogDir)
	mux.HandleFunc("/api/excel/convert", s.handleExcelConvert)
	mux.HandleFunc("/api/pe/read", s.handlePERead)
	mux.HandleFunc("/api/pe/convert", s.handlePEConvert)
	return mux
}

func writeJSON(w http.ResponseWriter, status int, v any) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.WriteHeader(status)
	_ = json.NewEncoder(w).Encode(v)
}

func readJSON(r *http.Request, v any) error {
	defer r.Body.Close()
	return json.NewDecoder(r.Body).Decode(v)
}

func (s *Server) handleConfig(w http.ResponseWriter, r *http.Request) {
	switch r.Method {
	case http.MethodGet:
		writeJSON(w, http.StatusOK, s.cfg)
	case http.MethodPost:
		var cfg config.Settings
		if err := readJSON(r, &cfg); err != nil {
			writeJSON(w, http.StatusBadRequest, map[string]string{"error": err.Error()})
			return
		}
		if cfg.UILang != "" {
			s.cfg.UILang = cfg.UILang
		}
		if cfg.OutputPath != "" {
			s.cfg.OutputPath = cfg.OutputPath
		}
		_ = config.Save(s.cfg)
		writeJSON(w, http.StatusOK, s.cfg)
	default:
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
	}
}

func (s *Server) handleI18n(w http.ResponseWriter, r *http.Request) {
	lang := strings.TrimPrefix(r.URL.Path, "/api/i18n/")
	if lang == "" {
		lang = "en"
	}
	writeJSON(w, http.StatusOK, i18n.All(lang))
}

func (s *Server) handleDialogFile(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}
	path, err := dialog.File().Filter("Excel files", "xlsx", "xls").Title("Select file").Load()
	if err != nil || path == "" {
		writeJSON(w, http.StatusOK, map[string]any{"path": ""})
		return
	}
	writeJSON(w, http.StatusOK, map[string]string{"path": path})
}

func (s *Server) handleDialogDir(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}
	var req struct {
		Initial string `json:"initial"`
	}
	_ = readJSON(r, &req)
	path, err := dialog.Directory().Browse()
	if err != nil || path == "" {
		writeJSON(w, http.StatusOK, map[string]any{"path": ""})
		return
	}
	writeJSON(w, http.StatusOK, map[string]string{"path": path})
}

func (s *Server) handleExcelConvert(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}
	var req struct {
		SourceFile string `json:"source_file"`
		OutputDir  string `json:"output_dir"`
	}
	if err := readJSON(r, &req); err != nil {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": err.Error()})
		return
	}
	if strings.TrimSpace(req.SourceFile) == "" {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": i18n.T("msg_no_file", s.cfg.UILang)})
		return
	}
	outputDir := req.OutputDir
	if outputDir == "" {
		outputDir = s.cfg.OutputPath
	}
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": err.Error()})
		return
	}
	tpl := template.DefaultPath()
	if err := template.Ensure(tpl); err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": i18n.T("msg_no_template", s.cfg.UILang) + ": " + err.Error()})
		return
	}
	out, err := convert.ProcessMenu(convert.ExportOptions{
		SourceFile:   req.SourceFile,
		TemplateFile: tpl,
		OutputDir:    outputDir,
	})
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": err.Error()})
		return
	}
	writeJSON(w, http.StatusOK, map[string]string{"output_file": out, "message": i18n.T("msg_done", s.cfg.UILang)})
}

func (s *Server) handlePERead(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}
	var req struct {
		SourceFile string `json:"source_file"`
	}
	if err := readJSON(r, &req); err != nil {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": err.Error()})
		return
	}
	if strings.TrimSpace(req.SourceFile) == "" {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": i18n.T("msg_no_file", s.cfg.UILang)})
		return
	}
	items, err := pe.ReadMenu(req.SourceFile)
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": err.Error()})
		return
	}
	s.peItems = items
	imgCount := 0
	for _, it := range items {
		if it.HasImage {
			imgCount++
		}
	}
	writeJSON(w, http.StatusOK, map[string]any{
		"items":       items,
		"item_count":  len(items),
		"image_count": imgCount,
	})
}

func (s *Server) handlePEConvert(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}
	var req struct {
		SourceFile string `json:"source_file"`
		OutputDir  string `json:"output_dir"`
	}
	if err := readJSON(r, &req); err != nil {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": err.Error()})
		return
	}
	if len(s.peItems) == 0 && strings.TrimSpace(req.SourceFile) != "" {
		items, err := pe.ReadMenu(req.SourceFile)
		if err != nil {
			writeJSON(w, http.StatusInternalServerError, map[string]string{"error": err.Error()})
			return
		}
		s.peItems = items
	}
	if len(s.peItems) == 0 {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": i18n.T("msg_no_file", s.cfg.UILang)})
		return
	}
	outputDir := req.OutputDir
	if outputDir == "" {
		outputDir = s.cfg.OutputPath
	}
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": err.Error()})
		return
	}
	tpl := template.DefaultPath()
	if err := template.Ensure(tpl); err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": i18n.T("msg_no_template", s.cfg.UILang) + ": " + err.Error()})
		return
	}
	out, err := convert.ProcessFromRows(pe.ToSourceRows(s.peItems), tpl, outputDir)
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": err.Error()})
		return
	}
	imgCount := 0
	if req.SourceFile != "" {
		imgCount, _ = pe.ExtractImages(req.SourceFile, s.peItems, outputDir)
	}
	picsDir := filepath.Join(outputDir, "pics")
	writeJSON(w, http.StatusOK, map[string]any{
		"output_file": out,
		"image_count": imgCount,
		"pics_dir":    picsDir,
		"message":     i18n.T("msg_pe_done", s.cfg.UILang),
	})
}
