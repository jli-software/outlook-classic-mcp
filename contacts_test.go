package main

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// peopleBox: ein eigener Kontakt und drei Eintraege im Firmen-Adressbuch,
// davon einer mit derselben Adresse wie der Kontakt.
func peopleBox() (*testBox, *fake) {
	b := newTestBox()
	contactsFolder := b.folder(b.root, "Kontakte", 2)
	b.dflt[olFolderContacts] = contactsFolder
	b.contact(contactsFolder, "Anna Muster", "anna@example.com", "Muster AG")
	b.person("Muster Anna", "anna@example.com", "IT")
	b.person("Muster Beat", "beat@example.com", "HR")
	b.person("Other Person", "other@example.com", "IT")
	return b, contactsFolder
}

func contactRun(t *testing.T, s *session, op string, p map[string]any) []contactInfo {
	t.Helper()
	if err := s.do(op, p); err != nil {
		t.Fatalf("%s: %v", op, err)
	}
	return parseJSONLines[contactInfo](s.out.Bytes())
}

func TestContactsSearch(t *testing.T) {
	b, _ := peopleBox()
	s := b.session()
	got := contactRun(t, s, "contacts_search", map[string]any{"query": "muster", "top": 10})
	if len(got) != 2 {
		t.Fatalf("got %d hits, want 2 (Verzeichnis-Dublette von Anna faellt weg): %+v", len(got), got)
	}
	if got[0].Name != "Anna Muster" || got[0].Source != "contacts" || got[0].Email != "anna@example.com" {
		t.Errorf("Kontakt: %+v", got[0])
	}
	// SMTP aus GetExchangeUser, nie der X.500-Name aus Address.
	if got[1].Name != "Muster Beat" || got[1].Email != "beat@example.com" || got[1].Department != "HR" || got[1].Source != "directory" {
		t.Errorf("Verzeichnis: %+v", got[1])
	}

	// Wortreihenfolge egal; der zweite Request kommt aus dem Cache.
	s2 := &session{ol: b.ol, ns: b.ns, dir: s.dir}
	got = contactRun(t, s2, "contacts_search", map[string]any{"query": "beat MUSTER"})
	if len(got) != 1 || got[0].Email != "beat@example.com" {
		t.Errorf("Wortreihenfolge: %+v", got)
	}
	if b.scans != 1 {
		t.Errorf("GAL %d-mal durchlaufen, want 1 (Cache)", b.scans)
	}
}

func TestContactsResolve(t *testing.T) {
	b, _ := peopleBox()
	b.known["beat@example.com"] = b.gal[1]

	// Outlook loest selbst auf.
	got := contactRun(t, b.session(), "contacts_resolve", map[string]any{"name": "beat@example.com"})
	if len(got) != 1 || !got[0].Resolved || got[0].Email != "beat@example.com" {
		t.Errorf("Resolve(): %+v", got)
	}

	// Outlook kann nicht, ein einziger Kandidat in anderer Wortreihenfolge.
	got = contactRun(t, b.session(), "contacts_resolve", map[string]any{"name": "Beat Muster"})
	if len(got) != 1 || !got[0].Resolved || got[0].Name != "Muster Beat" {
		t.Errorf("einziger Kandidat: %+v", got)
	}

	// Mehrdeutig: Liste, nichts aufgeloest.
	got = contactRun(t, b.session(), "contacts_resolve", map[string]any{"name": "Muster"})
	if len(got) != 2 {
		t.Fatalf("Kandidaten: %+v", got)
	}
	for _, c := range got {
		if c.Resolved {
			t.Errorf("geraten statt Liste: %+v", c)
		}
	}

	// Genau ein wortwoertlicher Treffer unter mehreren Kandidaten.
	b.person("Muster Beatrice", "beatrice@example.com", "HR")
	got = contactRun(t, b.session(), "contacts_resolve", map[string]any{"name": "muster beat"})
	if len(got) != 1 || !got[0].Resolved || got[0].Email != "beat@example.com" {
		t.Errorf("exakter Treffer: %+v", got)
	}

	if got = contactRun(t, b.session(), "contacts_resolve", map[string]any{"name": "zzz"}); len(got) != 0 {
		t.Errorf("kein Treffer: %+v", got)
	}
}

func TestContactsGetAndList(t *testing.T) {
	b, folder := peopleBox()
	z := b.contact(folder, "Zora Beispiel", "zora@example.com", "")
	cid := z.props["EntryID"].(string)
	gid := b.gal[1].props["ID"].(string)
	mail := b.mail(b.inbox, "keine Person", at(1)).props["EntryID"].(string)

	got := contactRun(t, b.session(), "contacts_get", map[string]any{"ids": []string{cid, gid, mail, "BAD"}})
	if len(got) != 4 {
		t.Fatalf("got %d, want 4: %+v", len(got), got)
	}
	if got[0].Name != "Zora Beispiel" || got[0].Notes != "Notiz zu Zora Beispiel" || got[0].Address != "Weg 2, 8000 Zurich" {
		t.Errorf("Kontakt-Details: %+v", got[0])
	}
	if got[1].Email != "beat@example.com" || got[1].Address != "Street 1, 8000 Zurich" {
		t.Errorf("Verzeichnis-Details: %+v", got[1])
	}
	if got[2].Error != "not a contact" || got[3].Error != "contact not found" {
		t.Errorf("Fehler je ID: %+v / %+v", got[2], got[3])
	}

	got = contactRun(t, b.session(), "contacts_list", map[string]any{"top": 1})
	if len(got) != 1 || got[0].Name != "Anna Muster" {
		t.Errorf("contacts_list sortiert/top: %+v", got)
	}
}

func TestWhoami(t *testing.T) {
	b := newTestBox()
	s := b.session()
	if err := s.do("whoami", nil); err != nil {
		t.Fatal(err) // CurrentUser des Namespace bricht ab — darf nicht durchschlagen
	}
	w := parseJSONLines[whoamiInfo](s.out.Bytes())
	if len(w) != 1 || w[0].Name != "Test User" || w[0].Email != "user@example.com" ||
		w[0].Connection != "cached" || w[0].OutlookVersion != "16.0.0.1" {
		t.Errorf("whoami: %+v", w)
	}
}

func TestReadShowsSenderSMTP(t *testing.T) {
	b := newTestBox()
	ext := b.mail(b.inbox, "extern", at(1))
	in := b.mail(b.inbox, "intern", at(2))
	in.props["SenderEmailType"] = "EX"
	in.props["SenderEmailAddress"] = "/O=EXAMPLE/CN=RECIPIENTS/CN=BEAT"
	in.props["Sender"] = b.person("Muster Beat", "beat@example.com", "HR")

	got := run(t, b.session(), "read", map[string]any{"ids": []string{ext.props["EntryID"].(string), in.props["EntryID"].(string)}})
	if got[0].FromEmail != "sender@example.com" || got[1].FromEmail != "beat@example.com" {
		t.Errorf("from_email: %q / %q", got[0].FromEmail, got[1].FromEmail)
	}
}

// ---- Pfad-Sandbox ----

func TestCheckPath(t *testing.T) {
	t.Setenv(allowAnyPathEnv, "")
	home, _ := os.UserHomeDir()
	for _, p := range []string{
		filepath.Join(os.TempDir(), "x.pdf"),
		filepath.Join(home, "Documents", "neu", "x.pdf"), // existiert nicht — trotzdem erlaubt
		home,
	} {
		if _, err := checkPath(p); err != nil {
			t.Errorf("%s: %v", p, err)
		}
	}
	for _, p := range []string{
		`C:\Windows\win.ini`,
		`\\server\share\x.pdf`,
		filepath.Join(home, "..", "x.pdf"),
		home + "x", // Nachbarordner mit gleichem Anfang
	} {
		if _, err := checkPath(p); err == nil || !strings.Contains(err.Error(), allowAnyPathEnv) {
			t.Errorf("%s: erlaubt (%v)", p, err)
		}
	}
	t.Setenv(allowAnyPathEnv, "1")
	if _, err := checkPath(`C:\Windows\win.ini`); err != nil {
		t.Errorf("abgeschaltet, trotzdem gesperrt: %v", err)
	}
}

func TestSafeFileName(t *testing.T) {
	for in, want := range map[string]string{
		"a.pdf": "a.pdf", `..\..\evil.exe`: "evil.exe", "../x/y.txt": "y.txt",
		"..": "attachment", "": "attachment", "C:evil.txt": "evil.txt", "a.txt:stream": "a.txt_stream",
	} {
		if got := safeFileName(in); got != want {
			t.Errorf("safeFileName(%q) = %q, want %q", in, got, want)
		}
	}
}

// Ein Pfad ausserhalb bricht den Antwort-Entwurf ab, bevor etwas entsteht.
func TestSandboxBlocksReplyAndSave(t *testing.T) {
	t.Setenv(allowAnyPathEnv, "")
	b := newTestBox()
	orig := b.mail(b.inbox, "Frage", at(1))
	id := orig.props["EntryID"].(string)
	before := len(b.byID)

	err := b.session().do("reply", map[string]any{"id": id, "body": "x", "attachments": []string{`C:\Windows\win.ini`}})
	if err == nil || !strings.Contains(err.Error(), "outside the allowed folders") {
		t.Errorf("reply: %v", err)
	}
	if len(b.byID) != before {
		t.Error("Entwurf trotz gesperrtem Pfad gespeichert")
	}
	err = b.session().do("save_attachment", map[string]any{"id": id, "dir": `C:\Windows\Temp\x`})
	if err == nil || !strings.Contains(err.Error(), "outside the allowed folders") {
		t.Errorf("save_attachment: %v", err)
	}
}
