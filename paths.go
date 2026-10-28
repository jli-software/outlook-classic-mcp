package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// Pfad-Sandbox: Dateien anhaengen (mail_reply) und Anhaenge speichern
// (mail_save_attachment) nur unter dem Benutzerprofil und %TEMP%. Grund:
// Eine praeparierte Mail kann per Prompt-Injection verlangen, beliebige
// Dateien anzuhaengen oder irgendwohin zu schreiben. Abschalten per
// OUTLOOK_MCP_ALLOW_ANY_PATH=1.
const allowAnyPathEnv = "OUTLOOK_MCP_ALLOW_ANY_PATH"

func sandboxOn() bool { return os.Getenv(allowAnyPathEnv) != "1" }

func allowedRoots() []string {
	var roots []string
	add := func(p string) {
		if p == "" {
			return
		}
		p = realPath(filepath.Clean(p))
		for _, r := range roots {
			if strings.EqualFold(r, p) {
				return
			}
		}
		roots = append(roots, p)
	}
	home, _ := os.UserHomeDir()
	add(home)
	add(os.TempDir())
	return roots
}

// checkPath liefert den absoluten Pfad oder einen Fehler, wenn er ausserhalb
// der erlaubten Wurzeln liegt. Symlinks und Junctions werden aufgeloest,
// soweit der Pfad existiert — sonst fuehrte ein Link im Profil hinaus.
func checkPath(p string) (string, error) {
	abs, err := filepath.Abs(p)
	if err != nil {
		return "", err
	}
	if !sandboxOn() {
		return abs, nil
	}
	real := realPath(abs)
	roots := allowedRoots()
	for _, r := range roots {
		if within(real, r) {
			return abs, nil
		}
	}
	return "", fmt.Errorf("path %s is outside the allowed folders (%s) — set %s=1 to allow any path",
		p, strings.Join(roots, ", "), allowAnyPathEnv)
}

// within: filepath.Rel vergleicht auf Windows ohne Gross-/Kleinschreibung und
// scheitert bei anderem Laufwerk oder UNC-Pfad.
func within(p, root string) bool {
	rel, err := filepath.Rel(root, p)
	if err != nil || filepath.IsAbs(rel) {
		return false
	}
	return rel != ".." && !strings.HasPrefix(rel, ".."+string(filepath.Separator))
}

// realPath loest Links im laengsten existierenden Teil des Pfads auf und
// haengt den Rest wieder an.
func realPath(p string) string {
	rest := ""
	cur := p
	for {
		if r, err := filepath.EvalSymlinks(cur); err == nil {
			return filepath.Join(r, rest)
		}
		parent := filepath.Dir(cur)
		if parent == cur {
			return p
		}
		rest = filepath.Join(filepath.Base(cur), rest)
		cur = parent
	}
}

// safeFileName: Der Dateiname eines Anhangs kommt vom Absender und darf
// nicht aus dem Zielordner hinausfuehren.
func safeFileName(name string) string {
	name = filepath.Base(strings.ReplaceAll(name, "/", `\`))
	name = strings.ReplaceAll(name, ":", "_") // "a.txt:x" waere ein NTFS-Datenstrom
	if name == "" || name == "." || name == ".." || name == `\` {
		return "attachment"
	}
	return name
}
