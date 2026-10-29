package main

import (
	"bufio"
	"context"
	"encoding/json"
	"os"
	"os/exec"
	"regexp"
	"runtime"
	"strings"
	"testing"
	"time"

	ole "github.com/go-ole/go-ole"
)

// Live-Test gegen das echte Postfach, ueber den echten MCP-stdio-Pfad — das ist
// der einzige Test, der die COM-Schicht wirklich anfasst. Laeuft nur mit
// gesetzter Umgebungsvariable, damit ein normales `go test` nicht ungefragt
// Outlook startet:
//
//	$env:OUTLOOK_MCP_LIVE = "1"; go test -run TestLive -v .
//
// Nur lesende Operationen. Kein mail_sync (das loest Senden/Empfangen aus und
// koennte im Postausgang haengende Mails rausschicken), kein Schreiben.
func requireLive(t *testing.T) {
	t.Helper()
	if os.Getenv("OUTLOOK_MCP_LIVE") == "" {
		t.Skip("OUTLOOK_MCP_LIVE nicht gesetzt — Live-Test gegen Outlook uebersprungen")
	}
}

type rpcClient struct {
	cmd    *exec.Cmd
	stdin  *json.Encoder
	stdout *bufio.Reader
	id     int
	t      *testing.T
}

func startServer(t *testing.T, args ...string) *rpcClient {
	t.Helper()
	return startServerExe(t, `.\outlook-classic-mcp.exe`, args...)
}

func startServerExe(t *testing.T, exe string, args ...string) *rpcClient {
	t.Helper()
	if _, err := os.Stat(exe); err != nil {
		t.Fatalf("%s fehlt — vorher `go build -o outlook-classic-mcp.exe .` ausfuehren", exe)
	}
	cmd := exec.Command(exe, append([]string{"serve"}, args...)...)
	in, err := cmd.StdinPipe()
	if err != nil {
		t.Fatal(err)
	}
	out, err := cmd.StdoutPipe()
	if err != nil {
		t.Fatal(err)
	}
	cmd.Stderr = os.Stderr
	if err := cmd.Start(); err != nil {
		t.Fatal(err)
	}
	c := &rpcClient{cmd: cmd, stdin: json.NewEncoder(in), stdout: bufio.NewReaderSize(out, 1<<20), t: t}
	t.Cleanup(func() {
		in.Close()
		done := make(chan struct{})
		go func() { _ = cmd.Wait(); close(done) }()
		select {
		case <-done:
		case <-time.After(10 * time.Second):
			_ = cmd.Process.Kill()
			<-done
		}
	})

	c.rpc("initialize", map[string]any{
		"protocolVersion": "2024-11-05",
		"capabilities":    map[string]any{},
		"clientInfo":      map[string]any{"name": "live-test", "version": "0"},
	})
	if err := c.stdin.Encode(map[string]any{
		"jsonrpc": "2.0", "method": "notifications/initialized",
	}); err != nil {
		t.Fatal(err)
	}
	return c
}

func (c *rpcClient) rpc(method string, params any) json.RawMessage {
	c.t.Helper()
	c.id++
	req := map[string]any{"jsonrpc": "2.0", "id": c.id, "method": method}
	if params != nil {
		req["params"] = params
	}
	if err := c.stdin.Encode(req); err != nil {
		c.t.Fatalf("%s: schreiben fehlgeschlagen: %v", method, err)
	}

	type resp struct {
		ID     int             `json:"id"`
		Result json.RawMessage `json:"result"`
		Error  *struct {
			Message string `json:"message"`
		} `json:"error"`
	}
	ch := make(chan resp, 1)
	errCh := make(chan error, 1)
	go func() {
		for {
			line, err := c.stdout.ReadString('\n')
			if err != nil {
				errCh <- err
				return
			}
			line = strings.TrimSpace(line)
			if line == "" {
				continue
			}
			var r resp
			if json.Unmarshal([]byte(line), &r) != nil || r.ID == 0 {
				continue // Notification o.ae.
			}
			ch <- r
			return
		}
	}()

	select {
	case r := <-ch:
		if r.Error != nil {
			c.t.Fatalf("%s: %s", method, r.Error.Message)
		}
		return r.Result
	case err := <-errCh:
		c.t.Fatalf("%s: Antwort abgebrochen: %v", method, err)
	case <-time.After(120 * time.Second):
		c.t.Fatalf("%s: Zeitueberschreitung", method)
	}
	return nil
}

// callTool ruft ein Tool und liefert den Text plus das isError-Flag.
func (c *rpcClient) callTool(name string, args map[string]any) (string, bool) {
	c.t.Helper()
	raw := c.rpc("tools/call", map[string]any{"name": name, "arguments": args})
	var res struct {
		Content []struct {
			Text string `json:"text"`
		} `json:"content"`
		IsError bool `json:"isError"`
	}
	if err := json.Unmarshal(raw, &res); err != nil {
		c.t.Fatalf("%s: Antwort nicht lesbar: %v", name, err)
	}
	var b strings.Builder
	for _, x := range res.Content {
		b.WriteString(x.Text)
	}
	return b.String(), res.IsError
}

type toolList struct {
	Tools []struct {
		Name        string `json:"name"`
		Description string `json:"description"`
		InputSchema struct {
			Properties map[string]json.RawMessage `json:"properties"`
			Required   []string                   `json:"required"`
		} `json:"inputSchema"`
	} `json:"tools"`
}

func listTools(t *testing.T, args ...string) toolList {
	t.Helper()
	c := startServer(t, args...)
	var res toolList
	if err := json.Unmarshal(c.rpc("tools/list", map[string]any{}), &res); err != nil {
		t.Fatal(err)
	}
	return res
}

// Senden ist opt-in: Mit --allow-send kommt genau mail_send_draft dazu, und
// die Entwurfs-Tools verweisen darauf.
func TestToolsListAllowSend(t *testing.T) {
	res := listTools(t, "--allow-send")
	if len(res.Tools) != 29 {
		t.Errorf("got %d Tools, want 29", len(res.Tools))
	}
	var send []string
	for _, tool := range res.Tools {
		if strings.Contains(tool.Name, "send") {
			send = append(send, tool.Name)
		}
		if tool.Name == "mail_reply" && !strings.Contains(tool.Description, "mail_send_draft") {
			t.Errorf("mail_reply verweist nicht auf mail_send_draft: %s", tool.Description)
		}
	}
	if strings.Join(send, ",") != "mail_send_draft" {
		t.Errorf("Sende-Tools: %v, want genau mail_send_draft", send)
	}
}

// Die Tool-Liste muss auch ohne Outlook stehen — reine Registrierung.
func TestToolsListWithoutOutlook(t *testing.T) {
	res := listTools(t)
	if len(res.Tools) != 28 {
		t.Errorf("got %d Tools, want 28", len(res.Tools))
	}

	byName := map[string]int{}
	for i, tool := range res.Tools {
		byName[tool.Name] = i
	}
	// Batch-Tools muessen 'id' UND 'ids' anbieten, und 'id' darf nicht mehr
	// required sein — sonst bricht der Batch-Aufruf.
	for _, name := range []string{"mail_read", "mail_move", "mail_mark_read",
		"mail_categorize", "mail_delete", "mail_attachments", "mail_save_attachment", "contacts_get",
		"calendar_get", "calendar_delete"} {
		i, ok := byName[name]
		if !ok {
			t.Errorf("%s fehlt", name)
			continue
		}
		props := res.Tools[i].InputSchema.Properties
		if _, ok := props["id"]; !ok {
			t.Errorf("%s: 'id' fehlt (alte Aufrufe wuerden brechen)", name)
		}
		if _, ok := props["ids"]; !ok {
			t.Errorf("%s: 'ids' fehlt", name)
		}
		for _, req := range res.Tools[i].InputSchema.Required {
			if req == "id" {
				t.Errorf("%s: 'id' ist required — Batch-Aufruf mit nur 'ids' bricht", name)
			}
		}
	}
	// mail_list/mail_search muessen include_body anbieten.
	for _, name := range []string{"mail_list", "mail_search"} {
		if _, ok := res.Tools[byName[name]].InputSchema.Properties["include_body"]; !ok {
			t.Errorf("%s: 'include_body' fehlt", name)
		}
	}
	// Ohne --allow-send bleibt Senden unmoeglich.
	for _, tool := range res.Tools {
		if strings.Contains(tool.Name, "send") {
			t.Errorf("Tool %q darf ohne --allow-send nicht existieren", tool.Name)
		}
		if strings.Contains(tool.Description, "mail_send_draft") {
			t.Errorf("%s verweist ohne --allow-send auf mail_send_draft", tool.Name)
		}
	}
}

// Fehlende IDs muessen als Tool-Fehler zurueckkommen, nicht als Absturz — und
// ohne Outlook anzufassen.
func TestMissingIDsWithoutOutlook(t *testing.T) {
	c := startServer(t)
	for _, name := range []string{"mail_read", "mail_delete", "mail_mark_read"} {
		text, isErr := c.callTool(name, map[string]any{})
		if !isErr {
			t.Errorf("%s ohne id/ids: kein Fehler, sondern %q", name, text)
		}
		if !strings.Contains(text, "'id'") {
			t.Errorf("%s: Fehlermeldung nennt den Parameter nicht: %q", name, text)
		}
	}
}

func TestLiveReadPath(t *testing.T) {
	requireLive(t)
	c := startServer(t)

	// 1. Ordnerbaum — erster Call, zieht den Worker- und COM-Start mit.
	start := time.Now()
	folders, isErr := c.callTool("mail_folders", map[string]any{})
	if isErr {
		t.Fatalf("mail_folders: %s", folders)
	}
	t.Logf("mail_folders: %v (erster Call, inkl. Worker+COM-Start)", time.Since(start).Round(time.Millisecond))
	if !strings.Contains(folders, "ID:") {
		t.Fatalf("keine Ordner-IDs in der Antwort: %.300s", folders)
	}

	// 2. Posteingang listen — Folge-Call, muss deutlich schneller sein.
	start = time.Now()
	list, isErr := c.callTool("mail_list", map[string]any{"top": 10})
	warmList := time.Since(start)
	if isErr {
		t.Fatalf("mail_list: %s", list)
	}
	t.Logf("mail_list(10): %v (Folge-Call)", warmList.Round(time.Millisecond))
	if strings.Contains(list, "No messages") {
		t.Skip("Posteingang leer — Rest des Live-Tests nicht moeglich")
	}

	ids := extractIDs(list)
	if len(ids) == 0 {
		t.Fatalf("keine IDs aus mail_list geparst: %.400s", list)
	}
	t.Logf("%d IDs gefunden", len(ids))

	// 3. Einzel-Read (alter Aufruf mit 'id') muss weiter funktionieren.
	start = time.Now()
	one, isErr := c.callTool("mail_read", map[string]any{"id": ids[0]})
	single := time.Since(start)
	if isErr {
		t.Fatalf("mail_read(id): %s", one)
	}
	if !strings.Contains(one, "Subject:") {
		t.Errorf("mail_read(id) ohne Betreff: %.200s", one)
	}
	t.Logf("mail_read(1 id): %v", single.Round(time.Millisecond))

	// 4. Der Kern: Batch-Read. Muss klar schneller sein als n Einzel-Reads.
	n := min(len(ids), 8)
	start = time.Now()
	many, isErr := c.callTool("mail_read", map[string]any{"ids": ids[:n]})
	batch := time.Since(start)
	if isErr {
		t.Fatalf("mail_read(ids): %s", many)
	}
	// Ueber die EntryIDs pruefen, nicht ueber "Subject:" — deutsche
	// Outlook-Zitate enthalten diesen String im Body und verfaelschen die Zaehlung.
	for i, id := range ids[:n] {
		if !strings.Contains(many, id) {
			t.Errorf("mail_read(%d ids): ID %d fehlt in der Antwort", n, i)
		}
	}
	if sep := strings.Count(many, strings.Repeat("─", 60)); sep != n-1 {
		t.Errorf("mail_read(%d ids): %d Trennlinien, want %d", n, sep, n-1)
	}
	// Dieselben n Mails einzeln lesen — nicht die erste Mail hochrechnen: die
	// Mails sind unterschiedlich gross, und eine kleine erste Mail liess den
	// Batch sonst langsamer aussehen, als er ist.
	var singles time.Duration
	for _, id := range ids[:n] {
		start = time.Now()
		c.callTool("mail_read", map[string]any{"id": id})
		singles += time.Since(start)
	}
	t.Logf("mail_read(%d ids): %v  —  %d Einzel-Reads: %v",
		n, batch.Round(time.Millisecond), n, singles.Round(time.Millisecond))
	// Ein Roundtrip kostet nur wenige ms; der Rest ist Mailgroesse und
	// Rauschen. Darum 50 % Spielraum — gefangen wird ein echter Rueckfall.
	if batch > singles*3/2 {
		t.Errorf("Batch (%v) ist deutlich langsamer als %d Einzel-Reads (%v)", batch, n, singles)
	}

	// 5. Ungueltige ID im Batch darf die gueltigen nicht mitreissen.
	mixed, isErr := c.callTool("mail_read", map[string]any{
		"ids": []string{ids[0], "0000UNGUELTIG0000"},
	})
	if isErr {
		t.Errorf("Batch mit einer kaputten ID komplett fehlgeschlagen: %.300s", mixed)
	}
	if !strings.Contains(mixed, "Subject:") {
		t.Errorf("gueltige ID im gemischten Batch verloren: %.300s", mixed)
	}
	if !strings.Contains(mixed, "✗") {
		t.Errorf("kaputte ID nicht als Fehler gemeldet: %.300s", mixed)
	}

	// 6. Unbekannter Ordner muss ein Fehler sein, nicht stillschweigend inbox.
	bad, isErr := c.callTool("mail_list", map[string]any{"folder": "GibtsGarantiertNicht_xyz42"})
	if !isErr {
		t.Errorf("unbekannter Ordner lieferte keinen Fehler: %.200s", bad)
	}

	// 7. include_body: Body direkt in der Liste, ohne zweiten Roundtrip.
	withBody, isErr := c.callTool("mail_list", map[string]any{"top": 2, "include_body": true})
	if isErr {
		t.Fatalf("mail_list(include_body): %s", withBody)
	}
	if !strings.Contains(withBody, "---") {
		t.Errorf("include_body lieferte keinen Body: %.300s", withBody)
	}

	// 8. Nach all dem muss der Worker noch leben.
	if again, isErr := c.callTool("mail_list", map[string]any{"top": 1}); isErr {
		t.Errorf("Worker nach der Testreihe unbrauchbar: %s", again)
	}
}

// extractIDs zieht die EntryIDs aus der formatierten mail_list-Ausgabe.
func extractIDs(list string) []string {
	var ids []string
	for _, line := range strings.Split(list, "\n") {
		line = strings.TrimSpace(line)
		if after, ok := strings.CutPrefix(line, "ID: "); ok {
			if id := strings.TrimSpace(after); id != "" {
				ids = append(ids, id)
			}
		}
	}
	return ids
}

// Suche live: ein Ordner und alle Ordner, mit Laufzeit. Suchbegriff ueber
// OUTLOOK_MCP_QUERY (Standard "the").
//
//	$env:OUTLOOK_MCP_LIVE = "1"; go test -run TestLiveSearch -v .
func TestLiveSearch(t *testing.T) {
	requireLive(t)
	query := os.Getenv("OUTLOOK_MCP_QUERY")
	if query == "" {
		query = "the"
	}
	c := startServer(t)
	for _, args := range []map[string]any{
		{"query": query, "top": 10},
		{"query": query, "folder": "all", "top": 10},
	} {
		start := time.Now()
		out, isErr := c.callTool("mail_search", args)
		if isErr {
			t.Fatalf("mail_search %v: %s", args, out)
		}
		t.Logf("mail_search %v: %v, %d hits", args, time.Since(start).Round(time.Millisecond), len(extractIDs(out)))
		if args["folder"] == "all" && len(extractIDs(out)) > 0 && !strings.Contains(out, "Folder: ") {
			t.Errorf("folder=all without folder per hit: %.300s", out)
		}
	}
}

// Schreibender Pfad, nur mit Entwuerfen: anlegen, aendern, Antwort-Entwurf,
// Anhaenge speichern — danach werden die Entwuerfe wieder geloescht (landen in
// "Gelöschte Elemente"). Nie gesendet.
//
//	$env:OUTLOOK_MCP_LIVE_WRITE = "1"; go test -run TestLiveWrite -v .
func TestLiveWriteDrafts(t *testing.T) {
	if os.Getenv("OUTLOOK_MCP_LIVE_WRITE") == "" {
		t.Skip("OUTLOOK_MCP_LIVE_WRITE nicht gesetzt — schreibender Live-Test uebersprungen")
	}
	c := startServer(t)
	idOf := func(text string) string {
		ids := extractIDs(text)
		if len(ids) == 0 {
			t.Fatalf("keine ID in %q", text)
		}
		return ids[0]
	}
	var drafts []string
	defer func() {
		if len(drafts) > 0 {
			out, isErr := c.callTool("mail_delete", map[string]any{"ids": drafts})
			t.Logf("Aufraeumen: %s (Fehler: %v)", out, isErr)
		}
	}()

	out, isErr := c.callTool("mail_create_draft", map[string]any{
		"subject": "outlook-classic-mcp live test – please ignore", "body": "Zeile 1\nGrüsse ✓", "to": []string{"test@example.invalid"},
	})
	if isErr {
		t.Fatalf("create_draft: %s", out)
	}
	draft := idOf(out)
	drafts = append(drafts, draft)

	if out, isErr = c.callTool("mail_update_draft", map[string]any{"id": draft, "body": "Neuer Text äöü", "cc": []string{}}); isErr {
		t.Fatalf("update_draft: %s", out)
	}
	read, _ := c.callTool("mail_read", map[string]any{"id": draft})
	if !strings.Contains(read, "Neuer Text äöü") || strings.Contains(read, "Zeile 1") {
		t.Errorf("update_draft nicht angekommen: %.400s", read)
	}

	list, _ := c.callTool("mail_list", map[string]any{"top": 1})
	orig := idOf(list)
	if out, isErr = c.callTool("mail_reply", map[string]any{"id": orig, "body": "Antwort vom Live-Test"}); isErr {
		t.Fatalf("reply: %s", out)
	}
	reply := idOf(out)
	drafts = append(drafts, reply)
	if out, isErr = c.callTool("mail_update_draft", map[string]any{"id": reply, "body": "Geaenderte Antwort"}); isErr {
		t.Fatalf("update_draft(reply): %s", out)
	}
	read, _ = c.callTool("mail_read", map[string]any{"id": reply})
	if !strings.Contains(read, "Geaenderte Antwort") || strings.Contains(read, "Antwort vom Live-Test") {
		t.Errorf("Antworttext nicht ersetzt: %.400s", read)
	}
	if !strings.Contains(read, "Von:") && !strings.Contains(read, "From:") {
		t.Errorf("Zitat verloren: %.600s", read)
	}

	// Anhaenge: erste Mail mit Anhang aus den letzten 25 speichern.
	list, _ = c.callTool("mail_list", map[string]any{"top": 25})
	atts, _ := c.callTool("mail_attachments", map[string]any{"ids": extractIDs(list)})
	t.Logf("Anhaenge:\n%.600s", atts)
	for _, id := range extractIDs(list) {
		a, _ := c.callTool("mail_attachments", map[string]any{"id": id})
		if strings.Contains(a, "no attachments") || strings.Contains(a, "✗") {
			continue
		}
		dir := t.TempDir()
		saved, isErr := c.callTool("mail_save_attachment", map[string]any{"id": id, "dir": dir})
		if isErr {
			t.Fatalf("save_attachment: %s", saved)
		}
		entries, _ := os.ReadDir(dir)
		if len(entries) == 0 {
			t.Errorf("nichts gespeichert: %s", saved)
		}
		t.Logf("gespeichert: %s", saved)
		break
	}
}

// Kontakte live, nur lesend: whoami, den eigenen Namen und die eigene Adresse
// aufloesen, einen Kollegen im Firmen-Adressbuch finden (Name ueber
// OUTLOOK_MCP_COLLEAGUE, sonst der eigene Nachname).
//
//	$env:OUTLOOK_MCP_LIVE = "1"; go test -run TestLiveContacts -v .
func TestLiveContacts(t *testing.T) {
	requireLive(t)
	c := startServer(t)

	who, isErr := c.callTool("outlook_whoami", map[string]any{})
	if isErr || !strings.Contains(who, "Mailbox: ") || !strings.Contains(who, "Sending: off") {
		t.Fatalf("outlook_whoami: %s", who)
	}
	line := strings.TrimPrefix(strings.SplitN(who, "\n", 2)[0], "Mailbox: ")
	name, email, _ := strings.Cut(strings.TrimSuffix(line, ">"), " <")
	if name == "" || !strings.Contains(email, "@") {
		t.Fatalf("whoami ohne Name/Adresse: %q", line)
	}

	for _, q := range []string{email, name} {
		start := time.Now()
		out, isErr := c.callTool("contacts_resolve", map[string]any{"name": q})
		t.Logf("contacts_resolve(<%d chars>): %v", len(q), time.Since(start).Round(time.Millisecond))
		if isErr || !strings.Contains(strings.ToLower(out), "resolved: ") || !strings.Contains(strings.ToLower(out), strings.ToLower(email)) {
			t.Errorf("eigener Name/Adresse nicht aufgeloest: %.300s", out)
		}
		if strings.Contains(out, "/O=") {
			t.Errorf("X.500-Adresse statt SMTP: %.300s", out)
		}
	}

	q := os.Getenv("OUTLOOK_MCP_COLLEAGUE")
	if q == "" {
		f := strings.Fields(name)
		q = f[len(f)-1]
	}
	for i := 0; i < 2; i++ { // zweiter Durchlauf aus dem Cache
		start := time.Now()
		out, isErr := c.callTool("contacts_search", map[string]any{"query": q, "top": 5})
		t.Logf("contacts_search #%d: %v, %d hits", i+1, time.Since(start).Round(time.Millisecond), len(extractIDs(out)))
		if isErr || len(extractIDs(out)) == 0 || !strings.Contains(out, "company directory") {
			t.Fatalf("contacts_search: %.400s", out)
		}
		if i == 1 && time.Since(start) > 2*time.Second {
			t.Errorf("zweite Suche nicht aus dem Cache: %v", time.Since(start))
		}
		if i == 1 {
			got, isErr := c.callTool("contacts_get", map[string]any{"id": extractIDs(out)[0]})
			if isErr || !strings.Contains(got, "@") {
				t.Errorf("contacts_get: %.300s", got)
			}
		}
	}

	if out, isErr := c.callTool("contacts_list", map[string]any{"top": 3}); isErr {
		t.Errorf("contacts_list: %s", out)
	}

	// Absender mit SMTP-Adresse.
	list, _ := c.callTool("mail_list", map[string]any{"top": 5})
	if ids := extractIDs(list); len(ids) > 0 {
		read, _ := c.callTool("mail_read", map[string]any{"ids": ids})
		if !strings.Contains(read, "@") || strings.Contains(read, "<\n") {
			t.Errorf("Absender ohne SMTP: %.300s", read)
		}
	}

	// Sandbox: ausserhalb des Profils wird abgelehnt.
	if ids := extractIDs(list); len(ids) > 0 {
		out, isErr := c.callTool("mail_save_attachment", map[string]any{"id": ids[0], "dir": `C:\Windows\Temp\ocm-test`})
		if !isErr || !strings.Contains(out, "outside the allowed folders") {
			t.Errorf("Sandbox greift nicht: %.300s", out)
		}
	}
}

// Eigene Kontakte live: legt einen Testkontakt an, findet ihn ueber den
// DASL-Filter und loescht ihn wieder (landet in "Gelöschte Elemente").
//
//	$env:OUTLOOK_MCP_LIVE_WRITE = "1"; go test -run TestLiveWriteContacts -v .
func TestLiveWriteContacts(t *testing.T) {
	if os.Getenv("OUTLOOK_MCP_LIVE_WRITE") == "" {
		t.Skip("OUTLOOK_MCP_LIVE_WRITE nicht gesetzt — schreibender Live-Test uebersprungen")
	}
	const tag = "Ocmtestkontakt"
	var id string
	_, err := outlook.call(context.Background(), "contacts_list", map[string]any{"top": 1}) // Verbindung aufbauen
	if err != nil {
		t.Fatal(err)
	}
	defer outlook.stop()
	err = liveCOM(func(s *session) error {
		c, err := callObj(s.ol, "CreateItem", 2) // olContactItem
		if err != nil {
			return err
		}
		c.put("FullName", "Zora "+tag)
		c.put("CompanyName", "Example Ltd")
		c.put("Email1Address", "zora@example.invalid")
		if _, err := c.call("Save"); err != nil {
			return err
		}
		id = str(field(c, "EntryID"))
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}
	defer liveCOM(func(s *session) error {
		c, err := callObj(s.ns, "GetItemFromID", id)
		if err == nil {
			_, err = c.call("Delete")
		}
		t.Logf("Aufraeumen: %v", err)
		return nil
	})
	for _, q := range []string{tag, "example.invalid", "example ltd " + strings.ToLower(tag)} {
		out, err := outlook.call(context.Background(), "contacts_search", map[string]any{"query": q, "top": 5})
		if err != nil {
			t.Fatal(err)
		}
		found := false
		for _, c := range parseJSONLines[contactInfo](out) {
			found = found || (c.ID == id && c.Email == "zora@example.invalid" && c.Source == "contacts")
		}
		if !found {
			t.Errorf("Testkontakt mit %q nicht gefunden: %s", q, out)
		}
	}
}

// liveCOM fuehrt fn auf einem eigenen STA-Thread mit eigener Verbindung aus —
// fuer Testaufbau, den kein Tool kann (Kontakt anlegen).
func liveCOM(fn func(*session) error) error {
	done := make(chan error, 1)
	go func() {
		runtime.LockOSThread()
		defer runtime.UnlockOSThread()
		ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
		defer ole.CoUninitialize()
		be, err := dialOutlook()
		if err != nil {
			done <- err
			return
		}
		defer be.close()
		ol, ns := be.begin()
		defer be.end()
		done <- fn(&session{ol: ol, ns: ns})
	}()
	return <-done
}

// Kalender live, nur lesend: diese Woche listen (prueft den Restrict-Filter
// mit ISO-Datum im echten Gebietsschema), Details holen, Frei/Gebucht der
// eigenen Adresse.
//
//	$env:OUTLOOK_MCP_LIVE = "1"; go test -run TestLiveCalendar -v .
func TestLiveCalendar(t *testing.T) {
	requireLive(t)
	c := startServer(t)

	start := time.Now()
	list, isErr := c.callTool("calendar_list", map[string]any{"top": 30})
	if isErr {
		t.Fatalf("calendar_list: %s", list)
	}
	ids := extractIDs(list)
	t.Logf("calendar_list (7 Tage): %v, %d Termine, davon %d Vorkommen", time.Since(start).Round(time.Millisecond),
		len(ids), strings.Count(list, "|20"))
	if len(ids) > 0 {
		got, isErr := c.callTool("calendar_get", map[string]any{"ids": ids[:min(len(ids), 5)]})
		if isErr || strings.Contains(got, "✗") {
			t.Errorf("calendar_get: %.600s", got)
		}
	}

	who, _ := c.callTool("outlook_whoami", map[string]any{})
	line := strings.TrimPrefix(strings.SplitN(who, "\n", 2)[0], "Mailbox: ")
	_, email, _ := strings.Cut(strings.TrimSuffix(line, ">"), " <")
	// Frei/Gebucht am Tag des ersten Termins mit Uhrzeit, zum Vergleich mit
	// dem Kalender (nur Zeiten ins Log, keine Betreffs).
	date := time.Now().Format("2006-01-02")
	if m := regexp.MustCompile(`\d{4}-\d\d-\d\d \d\d:\d\d`).FindString(list); m != "" {
		date = m[:10]
	}
	fb, isErr := c.callTool("calendar_free_busy", map[string]any{"emails": []string{email}, "date": date})
	if isErr || strings.Contains(fb, "✗") {
		t.Errorf("calendar_free_busy: %s", fb)
	}
	day, _ := c.callTool("calendar_list", map[string]any{"start": date, "end": date})
	var times []string
	for _, l := range strings.Split(day, "\n") {
		if m := regexp.MustCompile(`^\d+\. (.*?)  `).FindStringSubmatch(l); m != nil {
			times = append(times, m[1])
			if !strings.Contains(m[1], date) {
				t.Errorf("Termin ausserhalb des Tages: %s", m[1])
			}
		}
	}
	_, fbTimes, _ := strings.Cut(fb, "\n")
	t.Logf("%s — Frei/Gebucht:\n%sKalender: %v", date, fbTimes, times)
}

// Kalender schreibend: Termin anlegen, verschieben, loeschen; ungesendete
// Besprechung anlegen und loeschen; Serie (per COM angelegt) — ein
// Vorkommen verschieben, wiederfinden, loeschen. Alles am Testtag in einem
// Jahr, ohne Erinnerung, danach geloescht. Nie gesendet.
//
//	$env:OUTLOOK_MCP_LIVE_WRITE = "1"; go test -run TestLiveWriteCalendar -v .
func TestLiveWriteCalendar(t *testing.T) {
	if os.Getenv("OUTLOOK_MCP_LIVE_WRITE") == "" {
		t.Skip("OUTLOOK_MCP_LIVE_WRITE nicht gesetzt — schreibender Live-Test uebersprungen")
	}
	c := startServer(t)
	day := time.Now().AddDate(1, 0, 0)
	for day.Weekday() != time.Sunday {
		day = day.AddDate(0, 0, 1)
	}
	d := day.Format("2006-01-02")
	var created []string
	defer func() {
		for _, id := range created {
			out, isErr := c.callTool("calendar_delete", map[string]any{"id": id})
			t.Logf("Aufraeumen %s: %s (Fehler: %v)", id, strings.TrimSpace(out), isErr)
		}
	}()
	idOf := func(text string) string {
		ids := extractIDs(text)
		if len(ids) == 0 {
			t.Fatalf("keine ID in %q", text)
		}
		return ids[0]
	}

	out, isErr := c.callTool("calendar_create", map[string]any{
		"subject": "outlook-classic-mcp live test – please ignore", "start": d + "T06:00", "duration": 45,
		"location": "Nowhere", "body": "Grüsse ✓", "reminder": -1, "show_as": "free",
	})
	if isErr {
		t.Fatalf("calendar_create: %s", out)
	}
	appt := idOf(out)
	created = append(created, appt)

	out, isErr = c.callTool("calendar_update", map[string]any{"id": appt, "start": d + "T07:15"})
	if isErr || !strings.Contains(out, "07:15–08:00") {
		t.Errorf("calendar_update (Dauer bleibt): %s", out)
	}
	list, isErr := c.callTool("calendar_list", map[string]any{"start": d, "end": d})
	if isErr || !strings.Contains(list, appt) || !strings.Contains(list, "07:15–08:00") {
		t.Errorf("Termin nicht in calendar_list: %s", list)
	}
	got, _ := c.callTool("calendar_get", map[string]any{"id": appt})
	if !strings.Contains(got, "Grüsse ✓") || strings.Contains(got, "Reminder:") {
		t.Errorf("calendar_get: %s", got)
	}

	// Besprechung: gespeichert, nicht gesendet, darf geloescht werden.
	out, isErr = c.callTool("calendar_create", map[string]any{
		"subject": "outlook-classic-mcp live meeting – please ignore", "start": d + "T09:00",
		"attendees": []string{"test@example.invalid"}, "reminder": -1,
	})
	if isErr || !strings.Contains(out, "NOT sent") {
		t.Fatalf("calendar_create (Besprechung): %s", out)
	}
	meeting := idOf(out)
	out, isErr = c.callTool("calendar_delete", map[string]any{"id": meeting})
	if isErr || !strings.Contains(out, "✓ Deleted") {
		created = append(created, meeting)
		t.Errorf("ungesendete Besprechung nicht geloescht (FInvited?): %s", out)
	}

	// Serie: drei Tage, per COM angelegt (kein Tool legt Serien an).
	var series string
	err := liveCOM(func(s *session) error {
		a, err := callObj(s.ol, "CreateItem", olAppointmentItem)
		if err != nil {
			return err
		}
		a.put("Subject", "outlook-classic-mcp live series – please ignore")
		a.put("ReminderSet", false)
		a.put("BusyStatus", 0)
		pat, err := callObj(a, "GetRecurrencePattern")
		if err != nil {
			return err
		}
		pat.put("RecurrenceType", 0) // olRecursDaily
		pat.put("PatternStartDate", wall(day))
		pat.put("StartTime", time.Date(1899, 12, 30, 10, 0, 0, 0, time.UTC))
		pat.put("Duration", 30)
		pat.put("Occurrences", 3)
		if _, err := a.call("Save"); err != nil {
			return err
		}
		series = str(field(a, "EntryID"))
		return nil
	})
	if err != nil {
		t.Fatalf("Serie anlegen: %v", err)
	}
	created = append(created, series)
	end := day.AddDate(0, 0, 3).Format("2006-01-02")
	list, _ = c.callTool("calendar_list", map[string]any{"start": d, "end": end, "query": "live series"})
	occ := extractIDs(list)
	if len(occ) != 3 || !strings.HasPrefix(occ[1], series+"|") {
		t.Fatalf("Vorkommen: %v\n%s", occ, list)
	}
	out, isErr = c.callTool("calendar_update", map[string]any{"id": occ[1], "start": day.AddDate(0, 0, 1).Format("2006-01-02") + "T11:30"})
	if isErr {
		t.Fatalf("Vorkommen verschieben: %s", out)
	}
	moved := idOf(out)
	if !strings.HasSuffix(moved, "T11:30") {
		t.Errorf("verschobenes Vorkommen: %s", moved)
	}
	// Ausnahme: nur ueber den aktuellen Beginn auffindbar (Exceptions-Rueckfall).
	if got, isErr := c.callTool("calendar_get", map[string]any{"id": moved}); isErr || strings.Contains(got, "✗") {
		t.Errorf("Ausnahme nicht gefunden: %s", got)
	}
	if out, isErr = c.callTool("calendar_delete", map[string]any{"id": occ[2]}); isErr || strings.Contains(out, "✗") {
		t.Errorf("Vorkommen loeschen: %s", out)
	}
	list, _ = c.callTool("calendar_list", map[string]any{"start": d, "end": end, "query": "live series"})
	if n := len(extractIDs(list)); n != 2 || !strings.Contains(list, "11:30") {
		t.Errorf("nach Verschieben/Loeschen %d Vorkommen:\n%s", n, list)
	}
}

// Senden live — sendet WIRKLICH, darum hinter einer eigenen Variable: eine
// Mail an die eigene Adresse (nie an Dritte), Ablehnung eines unaufgeloesten
// Empfaengers, Weiterleitungs-Entwurf, Audit-Zeile. Danach werden alle Kopien
// geloescht (landen in "Gelöschte Elemente").
//
//	$env:OUTLOOK_MCP_LIVE_SEND = "1"; go test -run TestLiveSend -v .
func TestLiveSend(t *testing.T) {
	if os.Getenv("OUTLOOK_MCP_LIVE_SEND") == "" {
		t.Skip("OUTLOOK_MCP_LIVE_SEND nicht gesetzt — Live-Test mit echtem Senden uebersprungen")
	}
	c := startServer(t, "--allow-send")
	idOf := func(text string) string {
		ids := extractIDs(text)
		if len(ids) == 0 {
			t.Fatalf("keine ID in %q", text)
		}
		return ids[0]
	}
	var cleanup []string
	defer func() {
		if len(cleanup) > 0 {
			out, isErr := c.callTool("mail_delete", map[string]any{"ids": cleanup})
			t.Logf("Aufraeumen: %s (Fehler: %v)", strings.TrimSpace(out), isErr)
		}
	}()

	who, isErr := c.callTool("outlook_whoami", map[string]any{})
	if isErr || !strings.Contains(who, "Sending: on") {
		t.Fatalf("outlook_whoami: %s", who)
	}
	line := strings.TrimPrefix(strings.SplitN(who, "\n", 2)[0], "Mailbox: ")
	_, email, _ := strings.Cut(strings.TrimSuffix(line, ">"), " <")
	if !strings.Contains(email, "@") {
		t.Fatalf("eigene Adresse unbekannt: %q", line)
	}
	m := regexp.MustCompile(`Audit log: (.+)`).FindStringSubmatch(who)
	if m == nil {
		t.Fatalf("whoami ohne Audit-Log: %s", who)
	}
	auditFile := strings.TrimSpace(m[1])

	// Unaufgeloester Empfaenger: nichts geht raus.
	out, isErr := c.callTool("mail_create_draft", map[string]any{
		"subject": "outlook-classic-mcp unresolved test – please ignore", "body": "x", "to": []string{"Nobody Unresolvable Ocmtest"},
	})
	if isErr {
		t.Fatalf("create_draft: %s", out)
	}
	bad := idOf(out)
	cleanup = append(cleanup, bad)
	if out, isErr = c.callTool("mail_send_draft", map[string]any{"id": bad}); !isErr || !strings.Contains(out, "could not be resolved") {
		t.Errorf("unaufgeloester Empfaenger nicht abgelehnt: %s", out)
	}

	tag := "ocm-send-" + time.Now().Format("150405")
	subject := "outlook-classic-mcp send test " + tag + " – please ignore"
	out, isErr = c.callTool("mail_create_draft", map[string]any{"subject": subject, "body": "Grüsse ✓", "to": []string{email}})
	if isErr {
		t.Fatalf("create_draft: %s", out)
	}
	draft := idOf(out)
	start := time.Now()
	out, isErr = c.callTool("mail_send_draft", map[string]any{"id": draft})
	if isErr || !strings.Contains(strings.ToLower(out), strings.ToLower(email)) {
		cleanup = append(cleanup, draft)
		t.Fatalf("mail_send_draft: %s", out)
	}
	t.Logf("mail_send_draft: %v", time.Since(start).Round(time.Millisecond))

	data, _ := os.ReadFile(auditFile)
	if !strings.Contains(string(data), `"tool":"mail_send_draft"`) || !strings.Contains(string(data), tag) {
		t.Errorf("Senden nicht im Audit-Log %s", auditFile)
	}

	// Ankunft abwarten (ohne mail_sync — Outlook liefert selbst aus).
	var got []string
	for i := 0; i < 24 && len(got) < 2; i++ {
		time.Sleep(5 * time.Second)
		inbox, _ := c.callTool("mail_search", map[string]any{"query": tag, "top": 5})
		sent, _ := c.callTool("mail_search", map[string]any{"query": tag, "folder": "sentitems", "top": 5})
		got = append(extractIDs(inbox), extractIDs(sent)...)
	}
	cleanup = append(cleanup, got...)
	t.Logf("nach %v: %d Kopien (Posteingang + Gesendet)", time.Since(start).Round(time.Second), len(got))
	if len(got) < 2 {
		t.Errorf("Mail nicht angekommen oder nicht in Gesendet — Postausgang pruefen")
	}

	// Weiterleiten: nur Entwurf, an eine ungueltige Adresse, wird geloescht.
	list, _ := c.callTool("mail_list", map[string]any{"top": 1})
	out, isErr = c.callTool("mail_forward", map[string]any{"id": idOf(list), "to": []string{"test@example.invalid"}, "body": "FYI vom Live-Test"})
	if isErr {
		t.Fatalf("mail_forward: %s", out)
	}
	fw := idOf(out)
	cleanup = append(cleanup, fw)
	read, _ := c.callTool("mail_read", map[string]any{"id": fw})
	if !strings.Contains(read, "FYI vom Live-Test") || !strings.Contains(read, "test@example.invalid") {
		t.Errorf("Weiterleitung: %.400s", read)
	}
}
