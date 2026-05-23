package main

import (
	"context"
	"embed"
	"fmt"
	"io/fs"
	"log"
	"net"
	"net/http"
	"time"

	webview "github.com/webview/webview_go"

	"menu-converter-v4-go/internal/config"
	"menu-converter-v4-go/internal/i18n"
	"menu-converter-v4-go/internal/server"
)

//go:embed web/*
var webRoot embed.FS

const (
	windowWidth  = 960
	windowHeight = 720
	shutdownWait = 3 * time.Second
)

func main() {
	if err := run(); err != nil {
		log.Fatal(err)
	}
}

func run() error {
	webFS, err := fs.Sub(webRoot, "web")
	if err != nil {
		return fmt.Errorf("embed web assets: %w", err)
	}

	cfg := config.Load()
	srvApp := server.New(webFS, cfg)

	httpServer := &http.Server{
		Handler:      srvApp.Handler(),
		ReadTimeout:  30 * time.Second,
		WriteTimeout: 120 * time.Second,
		IdleTimeout:  60 * time.Second,
	}

	ln, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		return fmt.Errorf("listen: %w", err)
	}
	httpServer.Addr = ln.Addr().String()

	serveErr := make(chan error, 1)
	go func() {
		serveErr <- httpServer.Serve(ln)
	}()

	url := "http://" + httpServer.Addr
	w := webview.New(false)
	defer w.Destroy()

	w.SetTitle(i18n.T("title", cfg.UILang))
	w.SetSize(windowWidth, windowHeight, webview.HintNone)
	w.Navigate(url)

	// Run blocks until the user closes the window.
	w.Run()

	// Window closed — stop the background HTTP server and exit.
	if err := shutdownHTTPServer(httpServer); err != nil {
		return err
	}

	select {
	case err := <-serveErr:
		if err != nil && err != http.ErrServerClosed {
			return fmt.Errorf("http server: %w", err)
		}
	default:
	}

	return nil
}

func shutdownHTTPServer(srv *http.Server) error {
	ctx, cancel := context.WithTimeout(context.Background(), shutdownWait)
	defer cancel()

	if err := srv.Shutdown(ctx); err != nil {
		// Force-close lingering connections if graceful shutdown times out.
		_ = srv.Close()
		return fmt.Errorf("shutdown http server: %w", err)
	}
	return nil
}
