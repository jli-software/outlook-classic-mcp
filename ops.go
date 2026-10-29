package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"
)

// session fuehrt eine Op gegen Outlook aus. Die Ausgabe sind JSON Lines, eine
// Zeile pro Ergebnis; die typisierten Wrapper in outlook.go lesen sie ein.
type session struct {
	ol, ns obj
	out    bytes.Buffer
	dir    *dirCache // Firmen-Adressbuch, lebt laenger als der Request
}

// Outlook-Konstanten (OlDefaultFolders, OlBodyFormat, OlItemType).
const (
	olFolderDeletedItems = 3
	olFolderSentMail     = 5
	olFolderInbox        = 6
	olFolderDrafts       = 16
	olFolderSyncIssues   = 20
	olFolderJunk         = 23
	olFormatHTML         = 2
	olMailItem           = 0
)

func (s *session) emit(v map[string]any) {
	b, _ := json.Marshal(v)
	s.out.Write(b)
	s.out.WriteByte('\n')
}

func (s *session) do(op string, p map[string]any) error {
	switch op {
	case "list":
		return s.list(p)
	case "search":
		return s.search(p)
	case "read":
		return s.read(p)
	case "attachments":
		return s.attachments(p)
	case "save_attachment":
		return s.saveAttachment(p)
	case "draft":
		return s.draft(p)
	case "update_draft":
		return s.updateDraft(p)
	case "reply":
		return s.reply(p)
	case "forward":
		return s.forward(p)
	case "send_draft":
		return s.sendDraft(p)
	case "folders":
		return s.folders()
	case "create_folder":
		return s.createFolder(p)
	case "move":
		return s.move(p)
	case "mark":
		return s.mark(p)
	case "categorize":
		return s.categorize(p)
	case "delete":
		return s.delete(p)
	case "sync":
		return s.sync(p)
	case "contacts_search":
		return s.contactsSearch(p)
	case "contacts_resolve":
		return s.contactsResolve(p)
	case "contacts_get":
		return s.contactsGet(p)
	case "contacts_list":
		return s.contactsList(p)
	case "whoami":
		return s.whoami()
	case "calendar_list":
		return s.calendarList(p)
	case "calendar_get":
		return s.calendarGet(p)
	case "calendar_create":
		return s.calendarCreate(p)
	case "calendar_update":
		return s.calendarUpdate(p)
	case "calendar_delete":
		return s.calendarDelete(p)
	case "calendar_respond":
		return s.calendarRespond(p)
	case "calendar_free_busy":
		return s.calendarFreeBusy(p)
	}
	return fmt.Errorf("unknown operation: %s", op)
}

// ---- Lesen & Suchen ----

func (s *session) list(p map[string]any) error {
	folder, err := s.resolveFolder(pStr(p, "folder"))
	if err != nil {
		return err
	}
	items, err := getObj(folder, "Items")
	if err != nil {
		return err
	}
	items.call("Sort", "[ReceivedTime]", true)
	return s.emitTop(items, pInt(p, "top", 25), pBool(p, "include_body", false))
}

func (s *session) emitTop(items obj, top int, withBody bool) error {
	n := 0
	return eachItem(items, func(m obj) (bool, error) {
		if n >= top {
			return true, nil
		}
		s.emitItem(m, withBody, "")
		n++
		return false, nil
	})
}

func (s *session) search(p map[string]any) error {
	q := strings.ReplaceAll(pStr(p, "query"), "'", "''")
	filter := `@SQL=("urn:schemas:httpmail:subject" LIKE '%` + q + `%' OR "urn:schemas:httpmail:fromname" LIKE '%` + q +
		`%' OR "urn:schemas:httpmail:textdescription" LIKE '%` + q + `%')`
	top := pInt(p, "top", 25)
	withBody := pBool(p, "include_body", false)

	if strings.ToLower(pStr(p, "folder")) != "all" {
		folder, err := s.resolveFolder(pStr(p, "folder"))
		if err != nil {
			return err
		}
		all, err := getObj(folder, "Items")
		if err != nil {
			return err
		}
		items, err := callObj(all, "Restrict", filter)
		if err != nil {
			return err
		}
		items.call("Sort", "[ReceivedTime]", true)
		return s.emitTop(items, top, withBody)
	}

	root, err := s.rootFolder()
	if err != nil {
		return err
	}
	skip := ""
	if f, err := callObj(s.ns, "GetDefaultFolder", olFolderSyncIssues); err == nil {
		skip = str(field(f, "EntryID"))
	}
	var hits []searchHit
	err = eachSub(root, func(sf obj) (bool, error) {
		return false, s.searchFolders(sf, "", filter, top, &hits, skip)
	})
	if err != nil {
		return err
	}
	sort.SliceStable(hits, func(i, j int) bool { return hits[i].t.After(hits[j].t) })
	for i, h := range hits {
		if i >= top {
			break
		}
		s.emitItem(h.m, withBody, h.path)
	}
	return nil
}

type searchHit struct {
	t    time.Time
	m    obj
	path string
}

// searchFolders sammelt fuer folder=all die Treffer aller Mailordner
// (DefaultItemType 0, also keine Kalender, Kontakte, Aufgaben). Pro Ordner
// reichen die neuesten top Treffer, denn gesamthaft werden auch nur top
// ausgegeben.
// Laufzeit: Der LIKE-Scan ueber den Text kostet am echten Postfach ~7 s fuer
// ~2000 Mails. Items.Sort wertet den Filter erneut aus und hat das auf ~40 s
// getrieben, darum nur sortieren, wenn ein Ordner mehr als top Treffer hat
// (Falle 10). skip: Synchronisierungsprobleme, enthaelt nur Sync-Protokolle.
func (s *session) searchFolders(f obj, path, filter string, top int, hits *[]searchHit, skip string) error {
	if skip != "" && str(field(f, "EntryID")) == skip {
		return nil
	}
	p2 := path + `\` + str(field(f, "Name"))
	if kind, err := getInt(f, "DefaultItemType"); err != nil {
		if isTransient(err) {
			return err
		}
	} else if kind == olMailItem {
		err := func() error {
			all, err := getObj(f, "Items")
			if err != nil {
				return err
			}
			items, err := callObj(all, "Restrict", filter)
			if err != nil {
				return err
			}
			count, err := getInt(items, "Count")
			if err != nil || count == 0 {
				return err
			}
			if count > top {
				items.call("Sort", "[ReceivedTime]", true)
			}
			n := 0
			return eachItem(items, func(m obj) (bool, error) {
				if n >= top {
					return true, nil
				}
				t, _ := itemTime(m)
				*hits = append(*hits, searchHit{t, m, p2})
				n++
				return false, nil
			})
		}()
		// Ein Sonderordner, der Restrict nicht kann, soll die Suche nicht
		// abbrechen; ein transienter COM-Fehler wiederholt den ganzen Request.
		if err != nil && isTransient(err) {
			return err
		}
	}
	return eachSub(f, func(sf obj) (bool, error) {
		return false, s.searchFolders(sf, p2, filter, top, hits, skip)
	})
}

func (s *session) emitItem(m obj, withBody bool, folderPath string) {
	body := str(field(m, "Body"))
	preview := cutUTF16(strings.Join(strings.Fields(body), " "), 160)
	recv := ""
	if t, ok := itemTime(m); ok {
		recv = t.Format("2006-01-02 15:04")
	}
	o := map[string]any{
		"id":       str(field(m, "EntryID")),
		"subject":  str(field(m, "Subject")),
		"from":     str(field(m, "SenderName")),
		"received": recv,
		"preview":  preview,
		"unread":   field(m, "UnRead") == true,
	}
	if folderPath != "" {
		o["folder"] = folderPath
	}
	if withBody {
		o["body"] = body
	}
	s.emit(o)
}

// cutUTF16 kuerzt auf n UTF-16-Einheiten (so zaehlt Outlook selbst): ein
// Emoji zaehlt doppelt. Ein angebrochenes Surrogatpaar faellt weg.
func cutUTF16(s string, n int) string {
	units := 0
	for i, r := range s {
		w := 1
		if r >= 0x10000 {
			w = 2
		}
		if units+w > n {
			return s[:i]
		}
		units += w
	}
	return s
}

// itemTime: Empfangszeit, bei Entwuerfen u. ae. die letzte Aenderung.
// Outlook liefert Ortszeit ohne Zone; go-ole etikettiert sie als UTC, darum
// wird nur formatiert, nie umgerechnet.
func itemTime(m obj) (time.Time, bool) {
	if t, ok := field(m, "ReceivedTime").(time.Time); ok {
		return t, true
	}
	t, ok := field(m, "LastModificationTime").(time.Time)
	return t, ok
}

func (s *session) read(p map[string]any) error {
	return s.forEachID(p, func(id string) error {
		m, err := callObj(s.ns, "GetItemFromID", id)
		if err != nil {
			return err
		}
		recv := ""
		if t, ok := field(m, "ReceivedTime").(time.Time); ok {
			recv = t.Format("2006-01-02 15:04")
		}
		var atts []string
		if col, err := getObj(m, "Attachments"); err == nil {
			each(col, func(a obj) (bool, error) {
				atts = append(atts, str(field(a, "FileName")))
				return false, nil
			})
		}
		s.emit(map[string]any{
			"id":          str(field(m, "EntryID")),
			"subject":     str(field(m, "Subject")),
			"from":        str(field(m, "SenderName")),
			"from_email":  senderSMTP(m),
			"to":          str(field(m, "To")),
			"received":    recv,
			"body":        str(field(m, "Body")),
			"attachments": strings.Join(atts, ", "),
		})
		return nil
	})
}

// ---- Anhaenge ----

func (s *session) attachments(p map[string]any) error {
	return s.forEachID(p, func(id string) error {
		m, err := callObj(s.ns, "GetItemFromID", id)
		if err != nil {
			return err
		}
		subj := str(field(m, "Subject"))
		col, err := getObj(m, "Attachments")
		if err != nil {
			return err
		}
		found := false
		err = each(col, func(a obj) (bool, error) {
			s.emit(map[string]any{
				"id": id, "subject": subj,
				"index":    num(field(a, "Index")),
				"filename": str(field(a, "FileName")),
				"size":     num(field(a, "Size")),
				"type":     num(field(a, "Type")),
			})
			found = true
			return false, nil
		})
		if err != nil {
			return err
		}
		if !found {
			s.emit(map[string]any{"id": id, "subject": subj, "empty": true})
		}
		return nil
	})
}

func (s *session) saveAttachment(p map[string]any) error {
	dir := pStr(p, "dir")
	if dir == "" {
		dir = filepath.Join(os.TempDir(), "outlook-classic-mcp-attachments")
	}
	dir, err := checkPath(dir)
	if err != nil {
		return err
	}
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return err
	}
	only := pInt(p, "index", 0)
	// Zwei Mails im selben Batch koennen denselben Dateinamen haben — dann
	// darf die zweite die erste nicht ueberschreiben.
	used := map[string]bool{}
	return s.forEachID(p, func(id string) error {
		m, err := callObj(s.ns, "GetItemFromID", id)
		if err != nil {
			return err
		}
		subj := str(field(m, "Subject"))
		col, err := getObj(m, "Attachments")
		if err != nil {
			return err
		}
		found := false
		err = each(col, func(a obj) (bool, error) {
			if only > 0 && num(field(a, "Index")) != only {
				return false, nil
			}
			found = true
			name := safeFileName(str(field(a, "FileName")))
			path := filepath.Join(dir, name)
			if used[strings.ToLower(path)] {
				ext := filepath.Ext(name)
				base := strings.TrimSuffix(name, ext)
				i := 2
				for used[strings.ToLower(filepath.Join(dir, fmt.Sprintf("%s-%d%s", base, i, ext)))] {
					i++
				}
				path = filepath.Join(dir, fmt.Sprintf("%s-%d%s", base, i, ext))
			}
			used[strings.ToLower(path)] = true
			if _, err := a.call("SaveAsFile", path); err != nil {
				if isTransient(err) {
					return true, err
				}
				s.emit(map[string]any{"id": id, "subject": subj, "filename": name, "error": err.Error()})
				return false, nil
			}
			s.emit(map[string]any{"id": id, "subject": subj, "filename": name, "path": path, "size": num(field(a, "Size"))})
			return false, nil
		})
		if err != nil {
			return err
		}
		if !found {
			s.emit(map[string]any{"id": id, "subject": subj, "empty": true})
		}
		return nil
	})
}

// ---- Entwuerfe ----

func (s *session) draft(p map[string]any) error {
	m, err := callObj(s.ol, "CreateItem", olMailItem)
	if err != nil {
		return err
	}
	if err := m.put("Subject", pStr(p, "subject")); err != nil {
		return err
	}
	if err := m.put("Body", pStr(p, "body")); err != nil {
		return err
	}
	// Leere Felder nicht setzen: ein fehlendes cc darf keinen leeren
	// Empfaenger erzeugen.
	for _, f := range [][2]string{{"to", "To"}, {"cc", "CC"}, {"bcc", "BCC"}} {
		if v := strings.Join(pArr(p, f[0]), ";"); v != "" {
			if err := m.put(f[1], v); err != nil {
				return err
			}
		}
	}
	return s.saveDraft(m)
}

func (s *session) saveDraft(m obj) error {
	if r, err := getObj(m, "Recipients"); err == nil {
		r.call("ResolveAll")
	}
	if _, err := m.call("Save"); err != nil {
		return err
	}
	s.emit(map[string]any{"id": str(field(m, "EntryID")), "subject": str(field(m, "Subject"))})
	return nil
}

// updateDraft fasst nur die Felder an, deren Schluessel uebergeben wurde.
func (s *session) updateDraft(p map[string]any) error {
	m, err := callObj(s.ns, "GetItemFromID", pStr(p, "id"))
	if err != nil {
		return err
	}
	if pHas(p, "subject") {
		if err := m.put("Subject", pStr(p, "subject")); err != nil {
			return err
		}
	}
	if pHas(p, "body") {
		// Bei HTML nur den eigenen Text ersetzen (Falle 9).
		if format, _ := getInt(m, "BodyFormat"); format == olFormatHTML {
			html, err := getString(m, "HTMLBody")
			if err != nil {
				return err
			}
			err = m.put("HTMLBody", replaceDraftHTML(html, pStr(p, "body")))
			if err != nil {
				return err
			}
		} else if err := m.put("Body", pStr(p, "body")); err != nil {
			return err
		}
	}
	for _, f := range [][2]string{{"to", "To"}, {"cc", "CC"}, {"bcc", "BCC"}} {
		if pHas(p, f[0]) {
			if err := m.put(f[1], strings.Join(pArr(p, f[0]), ";")); err != nil {
				return err
			}
		}
	}
	return s.saveDraft(m)
}

func (s *session) reply(p map[string]any) error {
	// Pfade zuerst pruefen, bevor irgendetwas angelegt wird.
	var files []string
	for _, a := range pArr(p, "attachments") {
		if a == "" {
			continue
		}
		abs, err := checkPath(a)
		if err != nil {
			return err
		}
		files = append(files, abs)
	}
	orig, err := callObj(s.ns, "GetItemFromID", pStr(p, "id"))
	if err != nil {
		return err
	}
	verb := "Reply"
	if pBool(p, "reply_all", false) {
		verb = "ReplyAll"
	}
	r, err := callObj(orig, verb)
	if err != nil {
		return err
	}
	if err := prependText(r, pStr(p, "body")); err != nil {
		return err
	}
	rAtts, err := getObj(r, "Attachments")
	if err != nil {
		return err
	}
	if pBool(p, "keep_attachments", false) {
		if col, err := getObj(orig, "Attachments"); err == nil {
			each(col, func(a obj) (bool, error) {
				fp := filepath.Join(os.TempDir(), safeFileName(str(field(a, "FileName"))))
				if _, err := a.call("SaveAsFile", fp); err == nil {
					rAtts.call("Add", fp)
				}
				return false, nil
			})
		}
	}
	for _, a := range files {
		rAtts.call("Add", a)
	}
	if _, err := r.call("Save"); err != nil {
		return err
	}
	s.emit(map[string]any{"id": str(field(r, "EntryID")), "subject": str(field(r, "Subject"))})
	return nil
}

// prependText setzt den eigenen Text VOR das automatische Zitat einer Antwort
// oder Weiterleitung (Thread bleibt erhalten). Bei HTML nie Body setzen
// (Falle 9); nur Nur-Text-Mails bekommen den Text als Body.
func prependText(r obj, body string) error {
	if format, _ := getInt(r, "BodyFormat"); format == olFormatHTML {
		html, err := getString(r, "HTMLBody")
		if err != nil {
			return err
		}
		return r.put("HTMLBody", insertReplyHTML(html, body))
	}
	old, _ := getString(r, "Body")
	return r.put("Body", body+"\r\n\r\n"+old)
}

// forward legt einen Weiterleitungs-Entwurf an. Die Anhaenge des Originals
// nimmt Outlook selbst mit.
func (s *session) forward(p map[string]any) error {
	orig, err := callObj(s.ns, "GetItemFromID", pStr(p, "id"))
	if err != nil {
		return err
	}
	f, err := callObj(orig, "Forward")
	if err != nil {
		return err
	}
	if err := prependText(f, pStr(p, "body")); err != nil {
		return err
	}
	for _, x := range [][2]string{{"to", "To"}, {"cc", "CC"}, {"bcc", "BCC"}} {
		if v := strings.Join(pArr(p, x[0]), ";"); v != "" {
			if err := f.put(x[1], v); err != nil {
				return err
			}
		}
	}
	return s.saveDraft(f)
}

// ---- Senden (nur mit --allow-send) ----

// OlObjectClass der Elemente, die mail_send_draft aus "Entwuerfe" senden darf.
const (
	olMail                     = 43
	olMeetingResponseNegative  = 55
	olMeetingResponsePositive  = 56
	olMeetingResponseTentative = 57
	olCC                       = 2 // OlMailRecipientType; 1 = To
	olBCC                      = 3
)

// allowSend wird von `serve --allow-send` gesetzt. Ohne das Flag gibt es
// mail_send_draft gar nicht; die Op prueft es trotzdem (Falle 21).
var allowSend bool

// sendDraft verschickt einen Entwurf: nur aus dem Ordner Entwuerfe, nur Mails
// und Besprechungsantworten, nur mit lauter aufgeloesten Empfaengern. Alles
// wird VOR Send gelesen — danach liegt das Element im Postausgang und seine
// EntryID gilt nicht mehr. Ein Wiederholungsversuch nach Send scheitert an
// der Ordnerpruefung und kann darum nicht doppelt senden.
func (s *session) sendDraft(p map[string]any) error {
	if !allowSend {
		return fmt.Errorf("sending is disabled — the server was not started with --allow-send")
	}
	m, err := callObj(s.ns, "GetItemFromID", pStr(p, "id"))
	if err != nil {
		return err
	}
	drafts, err := callObj(s.ns, "GetDefaultFolder", olFolderDrafts)
	if err != nil {
		return err
	}
	parent, err := getObj(m, "Parent")
	if err != nil {
		return err
	}
	if !strings.EqualFold(str(field(parent, "EntryID")), str(field(drafts, "EntryID"))) {
		return fmt.Errorf("only items in the Drafts folder can be sent (this one is in %s)", str(field(parent, "Name")))
	}
	switch num(field(m, "Class")) {
	case olMail, olMeetingResponseNegative, olMeetingResponsePositive, olMeetingResponseTentative:
	default:
		return fmt.Errorf("only emails and meeting responses can be sent with this tool")
	}
	if field(m, "Sent") == true {
		return fmt.Errorf("this item was already sent")
	}
	rs, err := getObj(m, "Recipients")
	if err != nil {
		return err
	}
	if n, err := getInt(rs, "Count"); err != nil || n == 0 {
		if err != nil {
			return err
		}
		return fmt.Errorf("the draft has no recipients")
	}
	rs.call("ResolveAll")
	var to, cc, bcc, unresolved []string
	err = each(rs, func(r obj) (bool, error) {
		name := str(field(r, "Name"))
		if field(r, "Resolved") != true {
			unresolved = append(unresolved, name)
			return false, nil
		}
		addr := ""
		if ae, err := getObj(r, "AddressEntry"); err == nil {
			addr = smtpOf(ae)
		}
		if addr == "" {
			addr = str(field(r, "Address"))
		}
		if addr == "" || strings.HasPrefix(addr, "/") {
			addr = name
		}
		switch num(field(r, "Type")) {
		case olCC:
			cc = append(cc, addr)
		case olBCC:
			bcc = append(bcc, addr)
		default:
			to = append(to, addr)
		}
		return false, nil
	})
	if err != nil {
		return err
	}
	if len(unresolved) > 0 {
		return fmt.Errorf("not sent — these recipients could not be resolved: %s", strings.Join(unresolved, "; "))
	}
	id, subject := str(field(m, "EntryID")), str(field(m, "Subject"))
	if _, err := m.call("Send"); err != nil {
		return err
	}
	s.emit(map[string]any{"id": id, "subject": subject, "to": to, "cc": cc, "bcc": bcc, "sent": true})
	return nil
}

// ---- Ordner ----

func (s *session) rootFolder() (obj, error) {
	store, err := getObj(s.ns, "DefaultStore")
	if err != nil {
		return nil, err
	}
	return callObj(store, "GetRootFolder")
}

// resolveFolder loest einen Ordner auf: Alias | EntryID | Pfad ('\A\B' oder
// 'A/B') | blosser Name. Namen vergleicht es ohne Gross-/Kleinschreibung.
// Wirft einen Fehler, wenn nichts passt — KEIN stilles Zurueckfallen auf
// inbox, sonst liest der Agent die falschen Mails und merkt es nicht (Falle 7).
func (s *session) resolveFolder(spec string) (obj, error) {
	aliases := map[string]int{
		"": olFolderInbox, "inbox": olFolderInbox,
		"sentitems": olFolderSentMail, "sent": olFolderSentMail,
		"drafts":       olFolderDrafts,
		"deleteditems": olFolderDeletedItems, "trash": olFolderDeletedItems,
		"junk": olFolderJunk,
	}
	if n, ok := aliases[strings.ToLower(spec)]; ok {
		return callObj(s.ns, "GetDefaultFolder", n)
	}
	// EntryID?
	if f, err := callObj(s.ns, "GetFolderFromID", spec); err == nil {
		return f, nil
	} else if isTransient(err) {
		return nil, err
	}
	root, err := s.rootFolder()
	if err != nil {
		return nil, err
	}
	// Pfad?
	if strings.ContainsAny(spec, `\/`) {
		parts := strings.FieldsFunc(spec, func(r rune) bool { return r == '\\' || r == '/' })
		if len(parts) > 1 && strings.EqualFold(parts[0], str(field(root, "Name"))) {
			parts = parts[1:]
		}
		cur := root
		for _, part := range parts {
			var next obj
			err := eachSub(cur, func(sf obj) (bool, error) {
				if strings.EqualFold(str(field(sf, "Name")), part) {
					next = sf
					return true, nil
				}
				return false, nil
			})
			if err != nil {
				return nil, err
			}
			if next == nil {
				return nil, fmt.Errorf("folder not found (path): %s", spec)
			}
			cur = next
		}
		return cur, nil
	}
	// Blosser Name -> rekursiv suchen, muss eindeutig sein.
	var found []obj
	var walk func(f obj) error
	walk = func(f obj) error {
		return eachSub(f, func(sf obj) (bool, error) {
			if strings.EqualFold(str(field(sf, "Name")), spec) {
				found = append(found, sf)
			}
			return false, walk(sf)
		})
	}
	if err := walk(root); err != nil {
		return nil, err
	}
	switch len(found) {
	case 1:
		return found[0], nil
	case 0:
		return nil, fmt.Errorf("folder not found: %s", spec)
	}
	return nil, fmt.Errorf("folder name is ambiguous (%d matches): %s — use a path or EntryID", len(found), spec)
}

func (s *session) folders() error {
	root, err := s.rootFolder()
	if err != nil {
		return err
	}
	var walk func(f obj, path string) error
	walk = func(f obj, path string) error {
		p2 := path + `\` + str(field(f, "Name"))
		count := 0
		if items, err := getObj(f, "Items"); err == nil {
			count = num(field(items, "Count"))
		}
		s.emit(map[string]any{
			"id":     str(field(f, "EntryID")),
			"path":   p2,
			"name":   str(field(f, "Name")),
			"items":  count,
			"unread": num(field(f, "UnReadItemCount")),
		})
		return eachSub(f, func(sf obj) (bool, error) { return false, walk(sf, p2) })
	}
	return eachSub(root, func(sf obj) (bool, error) { return false, walk(sf, "") })
}

func (s *session) createFolder(p map[string]any) error {
	var parent obj
	var err error
	if id := pStr(p, "parent_id"); id != "" {
		parent, err = callObj(s.ns, "GetFolderFromID", id)
	} else {
		parent, err = callObj(s.ns, "GetDefaultFolder", olFolderInbox)
	}
	if err != nil {
		return err
	}
	fs, err := getObj(parent, "Folders")
	if err != nil {
		return err
	}
	nf, err := callObj(fs, "Add", pStr(p, "name"))
	if err != nil {
		return err
	}
	s.emit(map[string]any{"id": str(field(nf, "EntryID")), "path": str(field(nf, "FolderPath")), "name": str(field(nf, "Name"))})
	return nil
}

// ---- Verwaltung ----

func (s *session) move(p map[string]any) error {
	dest, err := callObj(s.ns, "GetFolderFromID", pStr(p, "folder_id"))
	if err != nil {
		return err
	}
	destPath := str(field(dest, "FolderPath"))
	return s.forEachID(p, func(id string) error {
		m, err := callObj(s.ns, "GetItemFromID", id)
		if err != nil {
			return err
		}
		moved, err := callObj(m, "Move", dest)
		if err != nil {
			return err
		}
		s.emit(map[string]any{"id": str(field(moved, "EntryID")), "subject": str(field(moved, "Subject")), "path": destPath})
		return nil
	})
}

func (s *session) mark(p map[string]any) error {
	read := pBool(p, "read", true)
	return s.forEachID(p, func(id string) error {
		m, err := callObj(s.ns, "GetItemFromID", id)
		if err != nil {
			return err
		}
		if err := m.put("UnRead", !read); err != nil {
			return err
		}
		if _, err := m.call("Save"); err != nil {
			return err
		}
		s.emit(map[string]any{"id": str(field(m, "EntryID")), "subject": str(field(m, "Subject")), "unread": field(m, "UnRead") == true})
		return nil
	})
}

func (s *session) categorize(p map[string]any) error {
	cats := strings.Join(pArr(p, "categories"), ", ")
	return s.forEachID(p, func(id string) error {
		m, err := callObj(s.ns, "GetItemFromID", id)
		if err != nil {
			return err
		}
		if err := m.put("Categories", cats); err != nil {
			return err
		}
		if _, err := m.call("Save"); err != nil {
			return err
		}
		s.emit(map[string]any{"id": str(field(m, "EntryID")), "subject": str(field(m, "Subject")), "categories": str(field(m, "Categories"))})
		return nil
	})
}

func (s *session) delete(p map[string]any) error {
	return s.forEachID(p, func(id string) error {
		m, err := callObj(s.ns, "GetItemFromID", id)
		if err != nil {
			return err
		}
		subj := str(field(m, "Subject"))
		if _, err := m.call("Delete"); err != nil {
			return err
		}
		s.emit(map[string]any{"id": id, "subject": subj, "deleted": true})
		return nil
	})
}

// sync stoesst Senden/Empfangen fuer alle Konten an und haelt so den
// OST-Cache frisch (wichtig, wenn das klassische Outlook selten laeuft).
func (s *session) sync(p map[string]any) error {
	if _, err := s.ns.call("SendAndReceive", false); err != nil {
		if so, err := getObj(s.ns, "SyncObjects"); err == nil {
			if one, err := callObj(so, "Item", 1); err == nil {
				one.call("Start")
			}
		}
	}
	time.Sleep(time.Duration(pInt(p, "wait", 8)) * time.Second)
	inbox, err := callObj(s.ns, "GetDefaultFolder", olFolderInbox)
	if err != nil {
		return err
	}
	count := 0
	if items, err := getObj(inbox, "Items"); err == nil {
		count = num(field(items, "Count"))
	}
	s.emit(map[string]any{"synced": true, "inbox_items": count, "inbox_unread": num(field(inbox, "UnReadItemCount"))})
	return nil
}

// forEachID: Ein fachlicher Fehler (ungueltige EntryID, Mail schon verschoben)
// betrifft nur diese eine Mail und wird als Objekt gemeldet, damit der Rest des
// Batches durchlaeuft. Ein transienter COM-Fehler muss den ganzen Request
// wiederholen und wird deshalb durchgereicht.
func (s *session) forEachID(p map[string]any, fn func(id string) error) error {
	var ids []string
	if id := pStr(p, "id"); id != "" {
		ids = append(ids, id)
	}
	for _, id := range pArr(p, "ids") {
		if id != "" {
			ids = append(ids, id)
		}
	}
	for _, id := range ids {
		if err := fn(id); err != nil {
			if isTransient(err) {
				return err
			}
			s.emit(map[string]any{"id": id, "error": err.Error()})
		}
	}
	return nil
}

// ---- Helfer fuer Objekte ----

func getObj(o obj, name string, args ...any) (obj, error) {
	v, err := o.get(name, args...)
	return asObj(name, v, err)
}

func callObj(o obj, name string, args ...any) (obj, error) {
	v, err := o.call(name, args...)
	return asObj(name, v, err)
}

func asObj(name string, v any, err error) (obj, error) {
	if err != nil {
		return nil, err
	}
	if r, ok := v.(obj); ok && r != nil {
		return r, nil
	}
	return nil, fmt.Errorf("%s returned no object", name)
}

func getString(o obj, name string) (string, error) {
	v, err := o.get(name)
	return str(v), err
}

func getInt(o obj, name string) (int, error) {
	v, err := o.get(name)
	return num(v), err
}

// field liest ein Property und schluckt Fehler — fuer Felder, die nicht jeder
// Elementtyp hat (Besprechungsanfragen, Berichte im Posteingang).
func field(o obj, name string) any {
	v, err := o.get(name)
	if err != nil {
		return nil
	}
	return v
}

func str(v any) string {
	switch x := v.(type) {
	case nil:
		return ""
	case string:
		return x
	}
	return fmt.Sprint(v)
}

func num(v any) int {
	switch x := v.(type) {
	case int:
		return x
	case float64:
		return int(x)
	case bool:
		if x {
			return 1
		}
	}
	return 0
}

// each laeuft ueber eine 1-basierte Collection (Folders, Attachments).
func each(col obj, fn func(obj) (bool, error)) error {
	n, err := getInt(col, "Count")
	if err != nil {
		return err
	}
	for i := 1; i <= n; i++ {
		it, err := callObj(col, "Item", i)
		if err != nil {
			return err
		}
		if stop, err := fn(it); stop || err != nil {
			return err
		}
	}
	return nil
}

func eachSub(f obj, fn func(obj) (bool, error)) error {
	fs, err := getObj(f, "Folders")
	if err != nil {
		return err
	}
	return each(fs, fn)
}

// eachItem laeuft per GetFirst/GetNext ueber Items. Das respektiert Sort und
// Restrict und ist bei grossen Ordnern schneller als Item(i).
func eachItem(items obj, fn func(obj) (bool, error)) error {
	v, err := items.call("GetFirst")
	for {
		if err != nil {
			return err
		}
		m, ok := v.(obj)
		if !ok || m == nil {
			return nil
		}
		if stop, err := fn(m); stop || err != nil {
			return err
		}
		v, err = items.call("GetNext")
	}
}

// ---- Helfer fuer Parameter ----

func pHas(p map[string]any, k string) bool {
	_, ok := p[k]
	return ok
}

func pStr(p map[string]any, k string) string { return str(p[k]) }

func pInt(p map[string]any, k string, d int) int {
	switch x := p[k].(type) {
	case int:
		return x
	case float64:
		return int(x)
	}
	return d
}

func pBool(p map[string]any, k string, d bool) bool {
	if b, ok := p[k].(bool); ok {
		return b
	}
	return d
}

func pArr(p map[string]any, k string) []string {
	switch x := p[k].(type) {
	case []string:
		return x
	case []any:
		out := make([]string, 0, len(x))
		for _, v := range x {
			out = append(out, str(v))
		}
		return out
	case string:
		if x != "" {
			return []string{x}
		}
	}
	return nil
}
