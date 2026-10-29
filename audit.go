package main

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"
)

// Audit-Log: jede schreibende Aktion als eine Zeile JSON — Zeit, Tool, IDs,
// Empfaenger, Betreff, Ergebnis. Nie Mailtext. Standardort
// %LOCALAPPDATA%\outlook-classic-mcp\audit.log; OUTLOOK_MCP_AUDIT_LOG setzt
// einen anderen Pfad oder schaltet es mit "off" ab. Ein Fehler beim Schreiben
// landet auf stderr und blockiert die Aktion nicht.
const (
	auditEnv     = "OUTLOOK_MCP_AUDIT_LOG"
	auditMaxSize = 5 << 20 // danach wird audit.log zu audit.log.1
)

var auditMu sync.Mutex

func auditPath() string {
	if v := strings.TrimSpace(os.Getenv(auditEnv)); v != "" {
		if strings.EqualFold(v, "off") {
			return ""
		}
		return v
	}
	base := os.Getenv("LOCALAPPDATA")
	if base == "" {
		base = os.TempDir()
	}
	return filepath.Join(base, "outlook-classic-mcp", "audit.log")
}

// audit schreibt eine Zeile. Leere Felder fallen weg; err != nil macht daraus
// einen Fehlschlag mit Meldung.
func audit(tool string, err error, fields map[string]any) {
	path := auditPath()
	if path == "" {
		return
	}
	rest := map[string]any{}
	for k, v := range fields {
		switch x := v.(type) {
		case nil:
			continue
		case string:
			if x == "" {
				continue
			}
		case []string:
			if len(x) == 0 {
				continue
			}
		case []map[string]any:
			if len(x) == 0 {
				continue
			}
		}
		rest[k] = v
	}
	rest["ok"] = err == nil
	if err != nil {
		rest["error"] = err.Error()
	}
	// Zeit und Tool vorne, der Rest alphabetisch (json.Marshal sortiert Maps).
	head, _ := json.Marshal(map[string]any{"time": time.Now().Format(time.RFC3339)})
	tl, _ := json.Marshal(tool)
	body, _ := json.Marshal(rest)
	line := string(head[:len(head)-1]) + `,"tool":` + string(tl) + "," + string(body[1:]) + "\n"

	auditMu.Lock()
	defer auditMu.Unlock()
	if werr := appendAudit(path, line); werr != nil {
		fmt.Fprintln(os.Stderr, "audit log:", werr)
	}
}

func appendAudit(path, line string) error {
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return err
	}
	if fi, err := os.Stat(path); err == nil && fi.Size()+int64(len(line)) > auditMaxSize {
		_ = os.Remove(path + ".1")
		if err := os.Rename(path, path+".1"); err != nil {
			return err
		}
	}
	f, err := os.OpenFile(path, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0o644)
	if err != nil {
		return err
	}
	if _, err := f.WriteString(line); err != nil {
		f.Close()
		return err
	}
	return f.Close()
}

// auditItems fasst Batch-Ergebnisse knapp zusammen: ID, Betreff, Fehler.
func auditItems(rs []actionResult) []map[string]any {
	var out []map[string]any
	for _, a := range rs {
		e := map[string]any{"id": a.ID, "subject": a.Subject}
		if a.Error != "" {
			e["error"] = a.Error
		}
		out = append(out, e)
	}
	return out
}

func auditCal(cs []calendarItem) []map[string]any {
	var out []map[string]any
	for _, c := range cs {
		e := map[string]any{"id": c.ID, "subject": c.Subject}
		if c.Error != "" {
			e["error"] = c.Error
		}
		out = append(out, e)
	}
	return out
}
