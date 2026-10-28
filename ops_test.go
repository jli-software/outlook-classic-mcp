package main

import (
	"context"
	"strings"
	"sync"
	"testing"
	"time"
)

func at(day int) time.Time { return time.Date(2026, 10, day, 9, 30, 0, 0, time.UTC) }

// testBox: Posteingang mit Unterordnern, Gesendet, Gelöschte, Kalender und
// Synchronisierungsprobleme — dazu zwei Ordner "Archiv" fuer Mehrdeutigkeit.
type testBox struct {
	*mailbox
	inbox, projects, deleted, calendar, syncIssues *fake
}

func newTestBox() *testBox {
	mb := newMailbox()
	b := &testBox{mailbox: mb}
	b.inbox = mb.folder(mb.root, "Posteingang", 0)
	b.projects = mb.folder(b.inbox, "Projects", 0)
	mb.folder(b.inbox, "Archiv", 0)
	mb.folder(mb.root, "Archiv", 0)
	b.deleted = mb.folder(mb.root, "Gelöschte Elemente", 0)
	b.calendar = mb.folder(mb.root, "Kalender", 1)
	b.syncIssues = mb.folder(mb.root, "Synchronisierungsprobleme", 0)
	mb.dflt[olFolderInbox] = b.inbox
	mb.dflt[olFolderDeletedItems] = b.deleted
	mb.dflt[olFolderSyncIssues] = b.syncIssues
	return b
}

func run(t *testing.T, s *session, op string, p map[string]any) []mailItem {
	t.Helper()
	if err := s.do(op, p); err != nil {
		t.Fatalf("%s: %v", op, err)
	}
	return parseJSONLines[mailItem](s.out.Bytes())
}

func TestResolveFolder(t *testing.T) {
	b := newTestBox()
	ok := []struct {
		spec string
		want *fake
	}{
		{"", b.inbox},
		{"INBOX", b.inbox},
		{"trash", b.deleted},
		{b.projects.props["EntryID"].(string), b.projects},
		{`\Posteingang\Projects`, b.projects},
		{`posteingang/projects`, b.projects},
		{`\user@example.com\Posteingang\Projects`, b.projects},
		{"projects", b.projects},
	}
	for _, tc := range ok {
		got, err := b.session().resolveFolder(tc.spec)
		if err != nil {
			t.Errorf("%q: %v", tc.spec, err)
		} else if got != tc.want {
			t.Errorf("%q: falscher Ordner %v", tc.spec, got.(*fake).props["Name"])
		}
	}

	bad := map[string]string{
		"Archiv":                 "ambiguous (2 matches)",
		"Rechnugen":              "folder not found: Rechnugen",
		`\Posteingang\Rechnugen`: "folder not found (path)",
	}
	for spec, want := range bad {
		if _, err := b.session().resolveFolder(spec); err == nil || !strings.Contains(err.Error(), want) {
			t.Errorf("%q: Fehler %v, erwartet %q — nie still auf den Posteingang zurueckfallen", spec, err, want)
		}
	}
}

func TestListNewestFirstWithTop(t *testing.T) {
	b := newTestBox()
	b.mail(b.inbox, "Mitte", at(2))
	b.mail(b.inbox, "Neu", at(3))
	b.mail(b.inbox, "Alt", at(1)).props["Body"] = "  Zeile 1\r\n\r\n\tZeile\u00a02  "

	got := run(t, b.session(), "list", map[string]any{"top": 2})
	if len(got) != 2 || got[0].Subject != "Neu" || got[1].Subject != "Mitte" {
		t.Fatalf("Reihenfolge/top falsch: %+v", got)
	}
	if got[0].Received != "2026-10-03 09:30" || !got[0].Unread || got[0].Body != "" {
		t.Errorf("Felder falsch: %+v", got[0])
	}

	got = run(t, b.session(), "list", map[string]any{"top": 5, "include_body": true})
	if got[2].Preview != "Zeile 1 Zeile 2" {
		t.Errorf("Vorschau nicht normalisiert: %q", got[2].Preview)
	}
	if !strings.Contains(got[2].Body, "\r\n") {
		t.Errorf("include_body liefert nicht den Originaltext: %q", got[2].Body)
	}
}

func TestPreviewCutsRunesNotBytes(t *testing.T) {
	b := newTestBox()
	b.mail(b.inbox, "Lang", at(1)).props["Body"] = strings.Repeat("ä", 200)
	got := run(t, b.session(), "list", map[string]any{"top": 1})
	if got[0].Preview != strings.Repeat("ä", 160) {
		t.Errorf("Vorschau: %d Zeichen, want 160", len([]rune(got[0].Preview)))
	}
}

func TestSearchAllFolders(t *testing.T) {
	b := newTestBox()
	b.mail(b.inbox, "Ticket 4711", at(1))
	b.mail(b.inbox, "anderes Thema", at(9))
	b.mail(b.projects, "Re: Ticket 4711", at(5))
	b.mail(b.deleted, "Fwd: Ticket 4711", at(3))
	b.mail(b.calendar, "Termin 4711", at(8))
	b.mail(b.syncIssues, "Sync 4711", at(7))

	got := run(t, b.session(), "search", map[string]any{"folder": "all", "query": "4711", "top": 2})
	if len(got) != 2 {
		t.Fatalf("top gilt gesamthaft: got %d Treffer: %+v", len(got), got)
	}
	if got[0].Subject != "Re: Ticket 4711" || got[0].Folder != `\Posteingang\Projects` {
		t.Errorf("neuester Treffer falsch: %+v", got[0])
	}
	if got[1].Subject != "Fwd: Ticket 4711" || got[1].Folder != `\Gelöschte Elemente` {
		t.Errorf("zweiter Treffer falsch (Gelöschte gehoeren dazu): %+v", got[1])
	}

	// Ein Ordner: kein Ordnerpfad im Ergebnis, Posteingang ohne Unterordner.
	got = run(t, b.session(), "search", map[string]any{"query": "4711", "top": 10})
	if len(got) != 1 || got[0].Folder != "" {
		t.Errorf("Suche im Posteingang: %+v", got)
	}
}

func TestSearchEscapesQuotes(t *testing.T) {
	b := newTestBox()
	var filter string
	b.inbox.props["Items"] = func() any {
		return fakeItems(nil).on("Restrict", func(args ...any) (any, error) {
			filter = args[0].(string)
			return fakeItems(nil), nil
		})
	}
	run(t, b.session(), "search", map[string]any{"query": "O'Brien"})
	if !strings.Contains(filter, "'%O''Brien%'") {
		t.Errorf("Apostroph nicht verdoppelt: %s", filter)
	}
}

// Ein Batch darf nicht komplett scheitern, weil eine ID ungueltig ist.
func TestReadBatchIsolatesPerIDErrors(t *testing.T) {
	b := newTestBox()
	a := b.mail(b.inbox, "Erste", at(1))
	c := b.mail(b.inbox, "Dritte", at(2))
	*a.props["_atts"].(*[]*fake) = []*fake{newFake(map[string]any{"FileName": "a.pdf"}), newFake(map[string]any{"FileName": "b.png"})}

	got := run(t, b.session(), "read", map[string]any{"ids": []string{a.props["EntryID"].(string), "BAD", c.props["EntryID"].(string)}})
	if len(got) != 3 {
		t.Fatalf("got %d Zeilen, want 3: %+v", len(got), got)
	}
	if got[0].Subject != "Erste" || got[0].Attachments != "a.pdf, b.png" || got[2].Subject != "Dritte" {
		t.Errorf("gute IDs nicht durchgelaufen: %+v", got)
	}
	if got[1].ID != "BAD" || got[1].Error != "Die Nachricht wurde nicht gefunden." {
		t.Errorf("kaputte ID ohne saubere Fehlermeldung: %+v", got[1])
	}
}

// Ein fehlendes cc darf keinen leeren Empfaenger setzen (frueher Falle 3).
func TestDraftSetsOnlyGivenRecipients(t *testing.T) {
	b := newTestBox()
	got := run(t, b.session(), "draft", map[string]any{"subject": "Hallo", "body": "Text", "to": []string{"a@b.ch", "c@d.ch"}})
	if len(got) != 1 || got[0].ID == "" || got[0].Subject != "Hallo" {
		t.Fatalf("Entwurf nicht gespeichert: %+v", got)
	}
	m := b.byID[got[0].ID]
	if m.props["To"] != "a@b.ch;c@d.ch" {
		t.Errorf("To: %v", m.props["To"])
	}
	if m.wasPut("CC") || m.wasPut("BCC") {
		t.Error("CC/BCC gesetzt, obwohl nicht uebergeben")
	}
}

func TestUpdateDraft(t *testing.T) {
	b := newTestBox()
	const quote = `<div><div style='border:none;border-top:solid #E1E1E1 1.0pt'><p>Von: X</p></div></div>`
	html := `<html><body><div class=WordSection1>` + textHTML("Alt") + quote + `</div></body></html>`
	m := b.mail(b.inbox, "AW: Frage", at(1))
	m.props["BodyFormat"] = olFormatHTML
	m.props["HTMLBody"] = html
	id := m.props["EntryID"].(string)

	run(t, b.session(), "update_draft", map[string]any{"id": id, "subject": "AW: Frage (neu)"})
	if m.wasPut("HTMLBody") || m.wasPut("Body") || m.wasPut("To") {
		t.Errorf("nur subject uebergeben, aber %v gesetzt", m.puts)
	}

	run(t, b.session(), "update_draft", map[string]any{"id": id, "body": "Neu", "cc": []string{}})
	if m.wasPut("Body") {
		t.Error("HTML-Entwurf auf Nur-Text umgestellt (Falle 9)")
	}
	if want := `<html><body><div class=WordSection1>` + textHTML("Neu") + quote + `</div></body></html>`; m.props["HTMLBody"] != want {
		t.Errorf("HTMLBody:\n got %s\nwant %s", m.props["HTMLBody"], want)
	}
	if m.props["CC"] != "" {
		t.Errorf("cc:[] muss CC leeren, ist %q", m.props["CC"])
	}

	plain := b.mail(b.inbox, "Notiz", at(2))
	run(t, b.session(), "update_draft", map[string]any{"id": plain.props["EntryID"].(string), "body": "Neu"})
	if plain.props["Body"] != "Neu" {
		t.Errorf("Nur-Text-Entwurf: Body %q", plain.props["Body"])
	}
}

func TestReplyKeepsHTMLAndAttaches(t *testing.T) {
	b := newTestBox()
	orig := b.mail(b.inbox, "Frage", at(1))
	got := run(t, b.session(), "reply", map[string]any{
		"id": orig.props["EntryID"].(string), "body": "Hallo", "attachments": []string{`C:\x\a.pdf`, ""},
	})
	if len(got) != 1 || got[0].Subject != "AW: Frage" {
		t.Fatalf("Antwort: %+v", got)
	}
	r := b.byID[got[0].ID]
	if want := `<html><body><div class=WordSection1>` + textHTML("Hallo") + `<p>Zitat</p></div></body></html>`; r.props["HTMLBody"] != want {
		t.Errorf("HTMLBody:\n got %s\nwant %s", r.props["HTMLBody"], want)
	}
	if r.wasPut("Body") {
		t.Error("Antwort auf Nur-Text umgestellt (Falle 9)")
	}
	if atts := *r.props["_atts"].(*[]*fake); len(atts) != 1 {
		t.Errorf("Anhaenge: %d, want 1 (leerer Pfad wird uebersprungen)", len(atts))
	}
}

func TestFoldersWalk(t *testing.T) {
	b := newTestBox()
	b.mail(b.projects, "x", at(1))
	s := b.session()
	if err := s.do("folders", nil); err != nil {
		t.Fatal(err)
	}
	fs := parseJSONLines[folderInfo](s.out.Bytes())
	var paths []string
	for _, f := range fs {
		paths = append(paths, f.Path)
		if f.Path == `\Posteingang\Projects` && f.Items != 1 {
			t.Errorf("Projects: %d Elemente, want 1", f.Items)
		}
	}
	want := `\Posteingang,\Posteingang\Projects,\Posteingang\Archiv,\Archiv,\Gelöschte Elemente,\Kalender,\Synchronisierungsprobleme`
	if strings.Join(paths, ",") != want {
		t.Errorf("Ordnerbaum:\n got %s\nwant %s", strings.Join(paths, ","), want)
	}
}

func TestMoveMarkCategorizeDelete(t *testing.T) {
	b := newTestBox()
	m := b.mail(b.inbox, "Rechnung", at(1))
	id := m.props["EntryID"].(string)
	dest := b.projects.props["EntryID"].(string)

	s := b.session()
	if err := s.do("move", map[string]any{"ids": []string{id}, "folder_id": dest}); err != nil {
		t.Fatal(err)
	}
	moved := parseJSONLines[actionResult](s.out.Bytes())
	if len(moved) != 1 || moved[0].ID == id || moved[0].Path != `\\user@example.com\Posteingang\Projects` {
		t.Fatalf("move: neue EntryID und Zielpfad erwartet: %+v", moved)
	}
	id = moved[0].ID

	s = b.session()
	s.do("mark", map[string]any{"ids": []string{id}, "read": true})
	s.do("categorize", map[string]any{"ids": []string{id}, "categories": []string{"Rot", "Projects"}})
	s.do("delete", map[string]any{"id": id})
	rs := parseJSONLines[actionResult](s.out.Bytes())
	if len(rs) != 3 || rs[0].Unread || rs[1].Categories != "Rot, Projects" || !rs[2].Deleted {
		t.Errorf("mark/categorize/delete: %+v", rs)
	}
}

func TestUnknownOp(t *testing.T) {
	if err := newTestBox().session().do("gibtsnicht", nil); err == nil || !strings.Contains(err.Error(), "unknown operation") {
		t.Errorf("got %v", err)
	}
}

// ---- STA-Thread ----

func newTestCOM(b *testBox, closed *int, dials *int) *outlookCOM {
	return &outlookCOM{
		dial:  func() (backend, error) { *dials++; return fakeBackend{b.mailbox, closed}, nil },
		sleep: func(time.Duration) {},
	}
}

// Transienter COM-Fehler (Outlook beschaeftigt): Verbindung verwerfen, ganzen
// Request wiederholen, Ausgabe des Fehlversuchs verwerfen.
func TestRetryOnTransientError(t *testing.T) {
	b := newTestBox()
	b.mail(b.inbox, "Hallo", at(1))
	fails := 2
	orig := b.ns.meths["GetDefaultFolder"]
	b.ns.on("GetDefaultFolder", func(args ...any) (any, error) {
		if fails > 0 {
			fails--
			return nil, &comError{0x80010001, "Der Aufruf wurde vom Aufgerufenen abgelehnt."}
		}
		return orig(args...)
	})
	var closed, dials int
	c := newTestCOM(b, &closed, &dials)
	defer c.stop()

	out, err := c.call(context.Background(), "list", map[string]any{"top": 5})
	if err != nil {
		t.Fatal(err)
	}
	if got := parseJSONLines[mailItem](out); len(got) != 1 {
		t.Errorf("Zeilen doppelt oder fehlend: %+v", got)
	}
	if dials != 3 || closed != 2 {
		t.Errorf("dials=%d closed=%d, want 3/2", dials, closed)
	}
}

// Fachlicher Fehler: kein Retry, Verbindung bleibt, Meldung kommt durch.
func TestNoRetryOnBusinessError(t *testing.T) {
	b := newTestBox()
	var closed, dials int
	c := newTestCOM(b, &closed, &dials)
	defer c.stop()

	_, err := c.call(context.Background(), "list", map[string]any{"folder": "Rechnugen"})
	if err == nil || err.Error() != "Outlook/COM error: folder not found: Rechnugen" {
		t.Fatalf("got %v", err)
	}
	if _, err := c.call(context.Background(), "folders", nil); err != nil {
		t.Fatalf("Folge-Call nach Fehler: %v", err)
	}
	if dials != 1 || closed != 0 {
		t.Errorf("dials=%d closed=%d, want 1/0", dials, closed)
	}
}

// Parallele Tool-Calls landen alle auf dem einen STA-Thread, nacheinander.
func TestConcurrentCallsSerialized(t *testing.T) {
	b := newTestBox()
	b.mail(b.inbox, "Hallo", at(1))
	var closed, dials int
	c := newTestCOM(b, &closed, &dials)
	var wg sync.WaitGroup
	errs := make(chan error, 20)
	for i := 0; i < 20; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			out, err := c.call(context.Background(), "list", map[string]any{"top": 5})
			if err == nil && len(parseJSONLines[mailItem](out)) != 1 {
				err = context.DeadlineExceeded
			}
			errs <- err
		}()
	}
	wg.Wait()
	close(errs)
	for err := range errs {
		if err != nil {
			t.Fatal(err)
		}
	}
	c.stop()
	if dials != 1 || closed != 1 {
		t.Errorf("dials=%d closed=%d, want 1/1 (stop gibt frei)", dials, closed)
	}
}

// .NET zaehlte UTF-16-Einheiten: ein Emoji belegt zwei.
func TestCutUTF16(t *testing.T) {
	for _, tc := range []struct {
		in   string
		n    int
		want string
	}{
		{"abc", 5, "abc"},
		{"abcdef", 3, "abc"},
		{"ab🤗cd", 4, "ab🤗"},
		{"ab🤗cd", 3, "ab"}, // Emoji haette nur halb Platz
		{"äöü", 2, "äö"},
	} {
		if got := cutUTF16(tc.in, tc.n); got != tc.want {
			t.Errorf("cutUTF16(%q, %d) = %q, want %q", tc.in, tc.n, got, tc.want)
		}
	}
}

// OLE-Datum auf Millisekunden wie DateTime.FromOADate, nicht auf Sekunden
// gerundet: 14:13:59,6 bleibt 14:13.
func TestFromOADate(t *testing.T) {
	day := 46323.0 // 28.10.2026
	frac := (14*3600 + 13*60 + 59.6) / 86400
	got := fromOADate(day + frac)
	if got.Format("2006-01-02 15:04:05.000") != "2026-10-28 14:13:59.600" {
		t.Errorf("got %s", got.Format("2006-01-02 15:04:05.000"))
	}
	if got := fromOADate(0); !got.Equal(time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)) {
		t.Errorf("Nullpunkt: %s", got)
	}
	// .NET: -1,25 = 29.12.1899 06:00 (Bruchteil zaehlt vorwaerts).
	if got := fromOADate(-1.25); got.Format("2006-01-02 15:04") != "1899-12-29 06:00" {
		t.Errorf("negativ: %s", got.Format("2006-01-02 15:04"))
	}
}
