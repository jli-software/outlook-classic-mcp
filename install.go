package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

const serverName = "outlook"

// runInstall traegt diesen Server automatisch in die Configs von Claude Desktop,
// Claude Code und Codex ein. Idempotent, mit Backup (.outlook-classic-mcp.bak).
func runInstall(args []string) {
	dryRun := false
	var only []string
	for _, a := range args {
		switch a {
		case "--print", "-n", "--dry-run":
			dryRun = true
		case "desktop", "code", "codex":
			only = append(only, a)
		default:
			fmt.Fprintf(os.Stderr, "Unknown argument %q (allowed: --print, desktop, code, codex)\n", a)
			os.Exit(2)
		}
	}

	exe, err := os.Executable()
	if err != nil {
		fmt.Fprintln(os.Stderr, "Cannot determine own path:", err)
		os.Exit(1)
	}
	exe = filepath.Clean(exe)

	home, _ := os.UserHomeDir()
	targets := []struct {
		key, label, path string
		toml             bool
	}{
		{"desktop", "Claude Desktop", claudeDesktopPath(), false},
		{"code", "Claude Code", filepath.Join(home, ".claude.json"), false},
		{"codex", "Codex", filepath.Join(home, ".codex", "config.toml"), true},
	}

	fmt.Printf("Server '%s' -> %s\n", serverName, exe)
	if dryRun {
		fmt.Println("(dry run — nothing is written)")
	}
	fmt.Println()

	for _, t := range targets {
		if len(only) > 0 && !contains(only, t.key) {
			continue
		}
		if dryRun {
			fmt.Printf("  %-15s %s\n", t.label, t.path)
			continue
		}
		var status string
		if t.toml {
			status, err = mergeCodexConfig(t.path, exe)
		} else {
			status, err = mergeJSONConfig(t.path, exe)
		}
		if err != nil {
			fmt.Printf("  ✗ %-15s %s\n      %v\n", t.label, t.path, err)
			continue
		}
		fmt.Printf("  ✓ %-15s %s  [%s]\n", t.label, t.path, status)
	}

	if !dryRun {
		fmt.Println("\nDone. Restart your clients so they load the new tools.")
	}
}

func contains(s []string, v string) bool {
	for _, x := range s {
		if x == v {
			return true
		}
	}
	return false
}

// claudeDesktopPath: Store-Version (virtualisiertes Roaming) bevorzugt, sonst Standard.
func claudeDesktopPath() string {
	matches, _ := filepath.Glob(filepath.Join(os.Getenv("LOCALAPPDATA"),
		"Packages", "Claude_*", "LocalCache", "Roaming", "Claude", "claude_desktop_config.json"))
	if len(matches) > 0 {
		return matches[0]
	}
	return filepath.Join(os.Getenv("APPDATA"), "Claude", "claude_desktop_config.json")
}

// mergeJSONConfig fuegt mcpServers.<serverName> in eine JSON-Config ein.
// Bestehende Inhalte bleiben erhalten (Round-Trip mit UseNumber).
func mergeJSONConfig(path, exe string) (string, error) {
	m := map[string]any{}
	if data, err := os.ReadFile(path); err == nil && len(bytes.TrimSpace(data)) > 0 {
		dec := json.NewDecoder(bytes.NewReader(data))
		dec.UseNumber()
		if err := dec.Decode(&m); err != nil {
			return "", fmt.Errorf("existing file is not valid JSON — left untouched: %w", err)
		}
	}

	servers, ok := m["mcpServers"].(map[string]any)
	if !ok || servers == nil {
		servers = map[string]any{}
	}
	_, existed := servers[serverName]
	servers[serverName] = map[string]any{
		"command": exe,
		"args":    []string{"serve"},
	}
	m["mcpServers"] = servers

	if err := backup(path); err != nil {
		return "", err
	}
	out, err := json.MarshalIndent(m, "", "  ")
	if err != nil {
		return "", err
	}
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return "", err
	}
	if err := os.WriteFile(path, out, 0o644); err != nil {
		return "", err
	}
	if existed {
		return "updated", nil
	}
	return "added", nil
}

// mergeCodexConfig setzt den [mcp_servers.<serverName>]-Block in der TOML.
// Existiert er schon, wird er (inkl. Pfad) ersetzt; sonst angehaengt. Andere
// Eintraege bleiben unangetastet.
func mergeCodexConfig(path, exe string) (string, error) {
	data, _ := os.ReadFile(path)
	content := string(data)
	header := "[mcp_servers." + serverName + "]"
	// Single-quoted TOML-Literal -> Backslashes im Windows-Pfad bleiben unangetastet.
	newBlock := fmt.Sprintf("%s\ncommand = '%s'\nargs = ['serve']", header, exe)

	if err := backup(path); err != nil {
		return "", err
	}

	if strings.Contains(content, header) {
		// Bestehenden Block (Header bis naechste [section] oder EOF) ersetzen.
		lines := strings.Split(content, "\n")
		var out []string
		for i := 0; i < len(lines); i++ {
			if strings.TrimSpace(lines[i]) == header {
				out = append(out, strings.Split(newBlock, "\n")...)
				i++
				for i < len(lines) && !strings.HasPrefix(strings.TrimSpace(lines[i]), "[") {
					i++
				}
				i-- // die naechste [section] wieder normal verarbeiten
				continue
			}
			out = append(out, lines[i])
		}
		if err := os.WriteFile(path, []byte(strings.Join(out, "\n")), 0o644); err != nil {
			return "", err
		}
		return "updated", nil
	}

	if len(content) > 0 && !strings.HasSuffix(content, "\n") {
		content += "\n"
	}
	content += "\n" + newBlock + "\n"
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return "", err
	}
	if err := os.WriteFile(path, []byte(content), 0o644); err != nil {
		return "", err
	}
	return "added", nil
}

// backup kopiert eine existierende Datei nach <path>.outlook-classic-mcp.bak.
func backup(path string) error {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil // existiert noch nicht -> kein Backup noetig
	}
	return os.WriteFile(path+".outlook-classic-mcp.bak", data, 0o644)
}
