package main

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/mark3labs/mcp-go/mcp"
)

// sendBox: Testpostfach mit Ordner "Entwuerfe" als Standardordner.
func sendBox() (*testBox, *fake) {
	b := newTestBox()
	drafts := b.folder(b.root, "Entwürfe", 0)
	b.dflt[olFolderDrafts] = drafts
	return b, drafts
}

func withSend(t *testing.T, on bool) {
	t.Helper()
	old := allowSend
	allowSend = on
	t.Cleanup(func() { allowSend = old })
}

func TestForwardCreatesHTMLDraft(t *testing.T) {
	b, drafts := sendBox()
	orig := b.mail(b.inbox, "Frage", at(1))
	got := run(t, b.session(), "forward", map[string]any{
		"id": orig.props["EntryID"].(string), "body": "Zur Info", "to": []string{"x@example.com"},
	})
	if len(got) != 1 || got[0].Subject != "WG: Frage" {
		t.Fatalf("Weiterleitung: %+v", got)
	}
	f := b.byID[got[0].ID]
	if want := `<html><body><div class=WordSection1>` + textHTML("Zur Info") + `<p>Zitat</p></div></body></html>`; f.props["HTMLBody"] != want {
		t.Errorf("HTMLBody:\n got %s\nwant %s", f.props["HTMLBody"], want)
	}
	if f.wasPut("Body") || f.wasPut("CC") || f.wasPut("BCC") {
		t.Errorf("unerwartet gesetzt: %v", f.puts)
	}
	if f.props["To"] != "x@example.com" || f.props["Parent"] != drafts || f.props["_sent"] == true {
		t.Errorf("To=%v Parent=%v gesendet=%v", f.props["To"], f.props["Parent"], f.props["_sent"])
	}
}

func TestSendDraft(t *testing.T) {
	withSend(t, true)
	b, _ := sendBox()
	got := run(t, b.session(), "draft", map[string]any{
		"subject": "Hallo", "body": "Text", "to": []string{"a@example.com"}, "cc": []string{"c@example.com"}, "bcc": []string{"d@example.com"},
	})
	m := b.byID[got[0].ID]
	s := b.session()
	if err := s.do("send_draft", map[string]any{"id": got[0].ID}); err != nil {
		t.Fatal(err)
	}
	sent := parseJSONLines[sentItem](s.out.Bytes())
	if len(sent) != 1 || !sent[0].Sent || sent[0].Subject != "Hallo" || sent[0].ID != got[0].ID ||
		strings.Join(sent[0].To, ",") != "a@example.com" || strings.Join(sent[0].CC, ",") != "c@example.com" ||
		strings.Join(sent[0].BCC, ",") != "d@example.com" {
		t.Errorf("Antwort: %+v", sent)
	}
	if m.props["_sent"] != true {
		t.Error("nicht gesendet")
	}
	// Ein zweiter Versuch (Retry nach Send) darf nicht nochmals senden.
	if err := b.session().do("send_draft", map[string]any{"id": got[0].ID}); err == nil {
		t.Error("zweites Senden ohne Fehler")
	}
}

// Jede Ablehnung muss vor Send greifen — der Fake protokolliert jedes Send.
func TestSendDraftRefuses(t *testing.T) {
	b, drafts := sendBox()
	draft := func(to string) *fake {
		m := b.mail(drafts, "Entwurf", at(1))
		m.props["To"] = to
		return m
	}
	inInbox := b.mail(b.inbox, "Eingang", at(1))
	inInbox.props["To"] = "a@example.com"
	appt := draft("a@example.com")
	appt.props["Class"] = 26 // olAppointment
	already := draft("a@example.com")
	already.props["Sent"] = true

	cases := []struct {
		name, want string
		m          *fake
		send       bool
	}{
		{"ohne --allow-send", "--allow-send", draft("a@example.com"), false},
		{"nicht in Entwuerfe", "only items in the Drafts folder", inInbox, true},
		{"falsche Klasse", "only emails and meeting responses", appt, true},
		{"schon gesendet", "already sent", already, true},
		{"keine Empfaenger", "no recipients", draft(""), true},
		{"unaufgeloest", "Anna Muster", draft("a@example.com;Anna Muster"), true},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			withSend(t, tc.send)
			err := b.session().do("send_draft", map[string]any{"id": tc.m.props["EntryID"].(string)})
			if err == nil || !strings.Contains(err.Error(), tc.want) {
				t.Errorf("got %v, want Fehler mit %q", err, tc.want)
			}
			if tc.m.props["_sent"] == true {
				t.Error("trotz Ablehnung gesendet")
			}
		})
	}
}

// Besprechungsantworten (calendar_respond) liegen in Entwuerfe und duerfen raus.
func TestSendDraftMeetingResponse(t *testing.T) {
	withSend(t, true)
	b, drafts := sendBox()
	m := b.mail(drafts, "Zugesagt: Sitzung", at(1))
	m.props["Class"] = olMeetingResponsePositive
	m.props["To"] = "org@example.com"
	if err := b.session().do("send_draft", map[string]any{"id": m.props["EntryID"].(string)}); err != nil {
		t.Fatal(err)
	}
	if m.props["_sent"] != true {
		t.Error("Antwort nicht gesendet")
	}
}

// Audit: eine Zeile JSON pro Aktion, Zeit und Tool vorne, nie Mailtext.
func TestAuditLog(t *testing.T) {
	path := filepath.Join(t.TempDir(), "sub", "audit.log")
	t.Setenv(auditEnv, path)
	withSend(t, true)
	b, _ := sendBox()
	var closed, dials int
	old := outlook
	outlook = newTestCOM(b, &closed, &dials)
	defer func() { outlook.stop(); outlook = old }()

	ctx := context.Background()
	res, _ := handleCreateDraft(ctx, newReq(map[string]any{
		"subject": "Hallo", "body": "GEHEIMER TEXT", "to": []any{"a@example.com"},
	}))
	id := extractIDs(resultText(res))[0]
	handleUpdateDraft(ctx, newReq(map[string]any{"id": id, "body": "NOCH GEHEIMER"}))
	handleSendDraft(ctx, newReq(map[string]any{"id": id}))
	handleSendDraft(ctx, newReq(map[string]any{"id": "GIBTSNICHT"}))

	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(data), "GEHEIM") {
		t.Errorf("Mailtext im Audit-Log:\n%s", data)
	}
	lines := strings.Split(strings.TrimSpace(string(data)), "\n")
	if len(lines) != 4 {
		t.Fatalf("%d Zeilen, want 4:\n%s", len(lines), data)
	}
	var tools []string
	for _, l := range lines {
		if !strings.HasPrefix(l, `{"time":"`) || !strings.Contains(l, `","tool":"`) {
			t.Errorf("Zeit/Tool nicht vorne: %s", l)
		}
		var e map[string]any
		if err := json.Unmarshal([]byte(l), &e); err != nil {
			t.Fatalf("kein JSON: %s", l)
		}
		tools = append(tools, e["tool"].(string))
	}
	if strings.Join(tools, ",") != "mail_create_draft,mail_update_draft,mail_send_draft,mail_send_draft" {
		t.Errorf("Tools: %v", tools)
	}
	var sent, failed map[string]any
	json.Unmarshal([]byte(lines[2]), &sent)
	json.Unmarshal([]byte(lines[3]), &failed)
	if sent["ok"] != true || sent["subject"] != "Hallo" || sent["id"] != id || !strings.Contains(lines[2], `"to":["a@example.com"]`) {
		t.Errorf("Senden: %s", lines[2])
	}
	if failed["ok"] != false || failed["error"] == nil {
		t.Errorf("Fehlschlag: %s", lines[3])
	}
	if !strings.Contains(lines[1], `"changed":["body"]`) {
		t.Errorf("update_draft: %s", lines[1])
	}

	t.Setenv(auditEnv, "off")
	handleSendDraft(ctx, newReq(map[string]any{"id": "GIBTSNICHT"}))
	if again, _ := os.ReadFile(path); len(again) != len(data) {
		t.Error("Audit-Log trotz off geschrieben")
	}
}

func resultText(r *mcp.CallToolResult) string {
	var b strings.Builder
	for _, c := range r.Content {
		if tc, ok := c.(mcp.TextContent); ok {
			b.WriteString(tc.Text)
		}
	}
	return b.String()
}

func TestAuditRotates(t *testing.T) {
	path := filepath.Join(t.TempDir(), "audit.log")
	t.Setenv(auditEnv, path)
	if err := os.WriteFile(path, make([]byte, auditMaxSize), 0o644); err != nil {
		t.Fatal(err)
	}
	audit("mail_delete", nil, nil)
	if fi, err := os.Stat(path + ".1"); err != nil || fi.Size() != auditMaxSize {
		t.Errorf("alte Datei nicht nach .1 verschoben: %v", err)
	}
	if data, _ := os.ReadFile(path); !strings.Contains(string(data), `"tool":"mail_delete","ok":true`) {
		t.Errorf("neue Datei: %q", data)
	}
}

func TestInstallArgs(t *testing.T) {
	dir := t.TempDir()
	js, toml := filepath.Join(dir, "c.json"), filepath.Join(dir, "config.toml")
	os.WriteFile(toml, []byte("[other]\nx = 1\n\n[mcp_servers.outlook]\ncommand = 'old'\nargs = ['serve']\n"), 0o644)
	for _, args := range [][]string{{"serve", "--allow-send"}, {"serve"}} {
		if _, err := mergeJSONConfig(js, `C:\x\o.exe`, args); err != nil {
			t.Fatal(err)
		}
		if _, err := mergeCodexConfig(toml, `C:\x\o.exe`, args); err != nil {
			t.Fatal(err)
		}
		var m struct {
			McpServers map[string]struct {
				Args []string `json:"args"`
			} `json:"mcpServers"`
		}
		data, _ := os.ReadFile(js)
		json.Unmarshal(data, &m)
		if got := strings.Join(m.McpServers[serverName].Args, " "); got != strings.Join(args, " ") {
			t.Errorf("JSON args %q, want %v", got, args)
		}
		tdata, _ := os.ReadFile(toml)
		want := "args = ['" + strings.Join(args, "', '") + "']"
		if !strings.Contains(string(tdata), want) || strings.Count(string(tdata), "[mcp_servers.outlook]") != 1 || !strings.Contains(string(tdata), "[other]") {
			t.Errorf("TOML:\n%s\nwant %s", tdata, want)
		}
	}
}
