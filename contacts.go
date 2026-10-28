package main

import (
	"fmt"
	"strings"
	"sync"
	"time"
)

// Kontakte, Firmen-Adressbuch (GAL) und das angemeldete Postfach.

const (
	olFolderContacts   = 10
	olContact          = 40 // OlObjectClass
	olDistributionList = 69

	// AddressEntryUserType
	olExchangeUserAddressEntry             = 0
	olExchangeDistributionListAddressEntry = 1
	olExchangeRemoteUserAddressEntry       = 5
	olOutlookContactAddressEntry           = 10

	prSMTPAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
)

// dirCache haelt Namen und IDs des Firmen-Adressbuchs zwischen den Requests.
// Ein Durchlauf kostet ~2 ms pro Eintrag (3600 Eintraege ~8 s), danach ist die
// Suche sofort da. Nur Strings, keine COM-Objekte — darum darf der Cache auch
// eine neue Verbindung oder einen neuen STA-Thread ueberleben.
type dirCache struct {
	mu    sync.Mutex
	at    time.Time
	names []string
	ids   []string
}

const (
	dirCacheTTL = 30 * time.Minute
	// dirScanBudget haelt einen Vollscan sicher unter callTimeout. Was bis
	// dahin nicht durchsucht ist, wird als unvollstaendig gemeldet.
	dirScanBudget = 45 * time.Second
)

func (c *dirCache) get() ([]string, []string, bool) {
	c.mu.Lock()
	defer c.mu.Unlock()
	if c.names == nil || time.Since(c.at) > dirCacheTTL {
		return nil, nil, false
	}
	return c.names, c.ids, true
}

func (c *dirCache) set(names, ids []string) {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.names, c.ids, c.at = names, ids, time.Now()
}

// matchWords: alle Woerter der Anfrage kommen im Text vor, Reihenfolge egal
// ("Anna Muster" findet "Muster Anna").
func matchWords(text string, words []string) bool {
	text = strings.ToLower(text)
	for _, w := range words {
		if !strings.Contains(text, w) {
			return false
		}
	}
	return true
}

// ---- Adressen ----

// smtpOf liefert die SMTP-Adresse eines AddressEntry. Exchange-Eintraege
// tragen in Address nur den X.500-Namen (/O=…/CN=…), der fuer eine Mail
// nichts taugt.
func smtpOf(ae obj) string {
	if ae == nil {
		return ""
	}
	switch num(field(ae, "AddressEntryUserType")) {
	case olExchangeUserAddressEntry, olExchangeRemoteUserAddressEntry:
		if xu, err := callObj(ae, "GetExchangeUser"); err == nil {
			if v := str(field(xu, "PrimarySmtpAddress")); v != "" {
				return v
			}
		}
	case olExchangeDistributionListAddressEntry:
		if dl, err := callObj(ae, "GetExchangeDistributionList"); err == nil {
			if v := str(field(dl, "PrimarySmtpAddress")); v != "" {
				return v
			}
		}
	}
	addr := str(field(ae, "Address"))
	if strings.EqualFold(str(field(ae, "Type")), "SMTP") {
		return addr
	}
	if pa, err := getObj(ae, "PropertyAccessor"); err == nil {
		if v, err := pa.call("GetProperty", prSMTPAddress); err == nil && str(v) != "" {
			return str(v)
		}
	}
	if strings.Contains(addr, "@") && !strings.HasPrefix(addr, "/") {
		return addr
	}
	return ""
}

// senderSMTP: Absender als SMTP-Adresse, auch bei internen Mails (Typ EX).
func senderSMTP(m obj) string {
	if !strings.EqualFold(str(field(m, "SenderEmailType")), "EX") {
		return str(field(m, "SenderEmailAddress"))
	}
	if ae, err := getObj(m, "Sender"); err == nil {
		return smtpOf(ae)
	}
	return ""
}

// contactEmail liest Email1..3 eines Kontakts; Exchange-Adressen ueber den
// zugehoerigen AddressEntry nach SMTP.
func (s *session) contactEmail(c obj, n int) string {
	key := fmt.Sprintf("Email%d", n)
	addr := str(field(c, key+"Address"))
	if !strings.EqualFold(str(field(c, key+"AddressType")), "EX") {
		return addr
	}
	if id := str(field(c, key+"EntryID")); id != "" {
		if ae, err := callObj(s.ns, "GetAddressEntryFromID", id); err == nil {
			if v := smtpOf(ae); v != "" {
				return v
			}
		}
	}
	return addr
}

// ---- Eintraege beschreiben ----

func (s *session) contactInfo(c obj, full bool) map[string]any {
	o := map[string]any{"id": str(field(c, "EntryID")), "source": "contacts"}
	if num(field(c, "Class")) == olDistributionList {
		o["name"] = str(field(c, "DLName"))
		o["kind"] = "list"
		return o
	}
	name := str(field(c, "FullName"))
	if name == "" {
		name = str(field(c, "FileAs"))
	}
	if name == "" {
		name = str(field(c, "Subject"))
	}
	o["name"] = name
	o["email"] = s.contactEmail(c, 1)
	for k, prop := range map[string]string{
		"company": "CompanyName", "department": "Department", "job_title": "JobTitle",
		"phone": "BusinessTelephoneNumber", "mobile": "MobileTelephoneNumber",
	} {
		o[k] = str(field(c, prop))
	}
	if full {
		o["email2"] = s.contactEmail(c, 2)
		o["email3"] = s.contactEmail(c, 3)
		o["home_phone"] = str(field(c, "HomeTelephoneNumber"))
		o["address"] = oneLine(str(field(c, "BusinessAddress")))
		o["notes"] = cutUTF16(strings.TrimSpace(str(field(c, "Body"))), 2000)
	}
	return o
}

func dirInfo(ae obj, full bool) map[string]any {
	o := map[string]any{
		"id": str(field(ae, "ID")), "name": str(field(ae, "Name")),
		"email": smtpOf(ae), "source": "directory",
	}
	switch num(field(ae, "AddressEntryUserType")) {
	case olExchangeDistributionListAddressEntry:
		o["kind"] = "list"
	case olExchangeUserAddressEntry, olExchangeRemoteUserAddressEntry:
		xu, err := callObj(ae, "GetExchangeUser")
		if err != nil {
			break
		}
		for k, prop := range map[string]string{
			"company": "CompanyName", "department": "Department", "job_title": "JobTitle",
			"office": "OfficeLocation", "phone": "BusinessTelephoneNumber", "mobile": "MobileTelephoneNumber",
		} {
			o[k] = str(field(xu, prop))
		}
		if full {
			addr := []string{str(field(xu, "StreetAddress")), strings.TrimSpace(str(field(xu, "PostalCode")) + " " + str(field(xu, "City")))}
			o["address"] = oneLine(strings.Join(addr, ", "))
			if mgr, err := callObj(xu, "GetExchangeUserManager"); err == nil {
				o["manager"] = str(field(mgr, "Name"))
			}
		}
	}
	return o
}

// oneLine macht aus einer mehrzeiligen Adresse eine Zeile.
func oneLine(s string) string {
	var parts []string
	for _, p := range strings.FieldsFunc(s, func(r rune) bool { return r == '\r' || r == '\n' }) {
		if p = strings.Trim(strings.TrimSpace(p), ","); p != "" {
			parts = append(parts, p)
		}
	}
	return strings.Join(parts, ", ")
}

// ---- Suchen ----

// findContacts sucht in den eigenen Kontakten (Name, Firma, E-Mail).
func (s *session) findContacts(words []string, top int) ([]map[string]any, error) {
	folder, err := callObj(s.ns, "GetDefaultFolder", olFolderContacts)
	if err != nil {
		if isTransient(err) {
			return nil, err
		}
		return nil, nil // Profil ohne Kontaktordner
	}
	all, err := getObj(folder, "Items")
	if err != nil {
		return nil, err
	}
	var conds []string
	for _, w := range words {
		w = strings.ReplaceAll(w, "'", "''")
		conds = append(conds, `("urn:schemas:contacts:cn" LIKE '%`+w+`%' OR "urn:schemas:contacts:fileas" LIKE '%`+w+
			`%' OR "urn:schemas:contacts:o" LIKE '%`+w+`%' OR "urn:schemas:contacts:email1" LIKE '%`+w+`%')`)
	}
	// Kann der Store den Filter nicht, selbst vergleichen — Kontaktordner
	// sind klein genug.
	items, err := callObj(all, "Restrict", "@SQL="+strings.Join(conds, " AND "))
	scan := false
	if err != nil {
		if isTransient(err) {
			return nil, err
		}
		items, scan = all, true
	}
	var out []map[string]any
	err = eachItem(items, func(c obj) (bool, error) {
		if len(out) >= top {
			return true, nil
		}
		if scan {
			text := str(field(c, "FullName")) + " " + str(field(c, "FileAs")) + " " +
				str(field(c, "CompanyName")) + " " + str(field(c, "Email1Address")) + " " + str(field(c, "DLName"))
			if !matchWords(text, words) {
				return false, nil
			}
		}
		out = append(out, s.contactInfo(c, false))
		return false, nil
	})
	return out, err
}

// findDirectory sucht im Firmen-Adressbuch nach Namen. Outlook kann die GAL
// nicht filtern, darum ein Durchlauf ueber alle Namen, danach aus dem Cache.
// note ist nicht leer, wenn nicht alles durchsucht wurde.
func (s *session) findDirectory(words []string, top int) (hits []map[string]any, note string, err error) {
	if s.dir == nil {
		s.dir = &dirCache{}
	}
	if names, ids, ok := s.dir.get(); ok {
		for i, n := range names {
			if len(hits) >= top {
				break
			}
			if !matchWords(n, words) {
				continue
			}
			ae, err := callObj(s.ns, "GetAddressEntryFromID", ids[i])
			if err != nil {
				if isTransient(err) {
					return nil, "", err
				}
				continue // inzwischen geloescht
			}
			hits = append(hits, dirInfo(ae, false))
		}
		return hits, "", nil
	}

	gal, err := callObj(s.ns, "GetGlobalAddressList")
	if err != nil {
		if isTransient(err) {
			return nil, "", err
		}
		return nil, "no company directory available", nil
	}
	ents, err := getObj(gal, "AddressEntries")
	if err != nil {
		return nil, "", err
	}
	start := time.Now()
	var names, ids []string
	v, err := ents.call("GetFirst")
	for {
		if err != nil {
			return nil, "", err
		}
		ae, ok := v.(obj)
		if !ok || ae == nil {
			break
		}
		name := str(field(ae, "Name"))
		names = append(names, name)
		ids = append(ids, str(field(ae, "ID")))
		if len(hits) < top && matchWords(name, words) {
			hits = append(hits, dirInfo(ae, false))
		}
		if time.Since(start) > dirScanBudget {
			total, _ := getInt(ents, "Count")
			return hits, fmt.Sprintf("company directory only partly searched (%d of %d entries)", len(names), total), nil
		}
		v, err = ents.call("GetNext")
	}
	s.dir.set(names, ids)
	return hits, "", nil
}

// findPeople: eigene Kontakte und Firmen-Adressbuch, je bis zu top Treffer.
// Ein Verzeichnistreffer mit derselben Adresse wie ein Kontakt faellt weg.
func (s *session) findPeople(query string, top int) ([]map[string]any, string, error) {
	words := strings.Fields(strings.ToLower(query))
	if len(words) == 0 {
		return nil, "", fmt.Errorf("query is empty")
	}
	contacts, err := s.findContacts(words, top)
	if err != nil {
		return nil, "", err
	}
	dir, note, err := s.findDirectory(words, top)
	if err != nil {
		return nil, "", err
	}
	seen := map[string]bool{}
	for _, c := range contacts {
		if e := strings.ToLower(str(c["email"])); e != "" {
			seen[e] = true
		}
	}
	out := contacts
	for _, d := range dir {
		if e := strings.ToLower(str(d["email"])); e == "" || !seen[e] {
			out = append(out, d)
		}
	}
	return out, note, nil
}

func (s *session) contactsSearch(p map[string]any) error {
	hits, note, err := s.findPeople(pStr(p, "query"), pInt(p, "top", 10))
	if err != nil {
		return err
	}
	for _, h := range hits {
		s.emit(h)
	}
	if note != "" {
		s.emit(map[string]any{"note": note})
	}
	return nil
}

// contactsResolve macht aus einem Namen eine SMTP-Adresse. Erst Outlooks
// eigene Aufloesung (schnell, eindeutig oder gar nicht), dann die Suche. Bei
// mehreren Kandidaten gibt es die Liste — nie raten.
func (s *session) contactsResolve(p map[string]any) error {
	name := strings.TrimSpace(pStr(p, "name"))
	if name == "" {
		return fmt.Errorf("name is empty")
	}
	r, err := callObj(s.ns, "CreateRecipient", name)
	if err != nil {
		return err
	}
	if ok, err := r.call("Resolve"); err != nil && isTransient(err) {
		return err
	} else if ok == true {
		if o := s.recipientInfo(r); str(o["email"]) != "" {
			o["resolved"] = true
			s.emit(o)
			return nil
		}
	}

	hits, note, err := s.findPeople(name, 10)
	if err != nil {
		return err
	}
	// Genau ein Kandidat, oder genau einer, der wortwoertlich passt
	// (Name in beliebiger Wortreihenfolge oder die Adresse selbst).
	want := sortedWords(name)
	var exact []map[string]any
	for _, h := range hits {
		if sortedWords(str(h["name"])) == want || strings.EqualFold(str(h["email"]), name) {
			exact = append(exact, h)
		}
	}
	switch {
	case len(hits) == 1:
		hits[0]["resolved"] = true
	case len(exact) == 1:
		exact[0]["resolved"] = true
		hits = exact
	}
	for _, h := range hits {
		s.emit(h)
	}
	if note != "" {
		s.emit(map[string]any{"note": note})
	}
	return nil
}

func sortedWords(s string) string {
	w := strings.Fields(strings.ToLower(s))
	for i := 1; i < len(w); i++ {
		for j := i; j > 0 && w[j] < w[j-1]; j-- {
			w[j], w[j-1] = w[j-1], w[j]
		}
	}
	return strings.Join(w, " ")
}

// recipientInfo beschreibt einen aufgeloesten Recipient je nach Herkunft.
func (s *session) recipientInfo(r obj) map[string]any {
	ae, err := getObj(r, "AddressEntry")
	if err != nil {
		return map[string]any{"name": str(field(r, "Name")), "email": str(field(r, "Address"))}
	}
	switch num(field(ae, "AddressEntryUserType")) {
	case olExchangeUserAddressEntry, olExchangeDistributionListAddressEntry, olExchangeRemoteUserAddressEntry:
		return dirInfo(ae, false)
	case olOutlookContactAddressEntry:
		if c, err := callObj(ae, "GetContact"); err == nil {
			return s.contactInfo(c, false)
		}
	}
	return map[string]any{"name": str(field(r, "Name")), "email": smtpOf(ae), "source": "address"}
}

// contactsGet: eine ID ist entweder ein Kontakt (EntryID) oder ein Eintrag
// im Firmen-Adressbuch (AddressEntry.ID).
func (s *session) contactsGet(p map[string]any) error {
	return s.forEachID(p, func(id string) error {
		if c, err := callObj(s.ns, "GetItemFromID", id); err == nil {
			if k := num(field(c, "Class")); k != olContact && k != olDistributionList {
				return fmt.Errorf("not a contact")
			}
			s.emit(s.contactInfo(c, true))
			return nil
		} else if isTransient(err) {
			return err
		}
		ae, err := callObj(s.ns, "GetAddressEntryFromID", id)
		if err != nil {
			if isTransient(err) {
				return err
			}
			return fmt.Errorf("contact not found")
		}
		s.emit(dirInfo(ae, true))
		return nil
	})
}

func (s *session) contactsList(p map[string]any) error {
	folder, err := callObj(s.ns, "GetDefaultFolder", olFolderContacts)
	if err != nil {
		return err
	}
	items, err := getObj(folder, "Items")
	if err != nil {
		return err
	}
	items.call("Sort", "[Subject]", false)
	top, n := pInt(p, "top", 50), 0
	return eachItem(items, func(c obj) (bool, error) {
		if n >= top {
			return true, nil
		}
		s.emit(s.contactInfo(c, false))
		n++
		return false, nil
	})
}

// ---- Wer bin ich ----

var connectionModes = map[int]string{
	0: "no Exchange", 100: "offline", 200: "cached, offline", 300: "disconnected",
	400: "cached, disconnected", 500: "cached, headers only", 600: "cached, downloading",
	700: "cached", 800: "online",
}

// whoami: Namespace.CurrentUser bricht kurz nach dem Start von Outlook mit
// E_ABORT ab; ueber das Konto klappt es sofort, darum zuerst Accounts.
func (s *session) whoami() error {
	o := map[string]any{"outlook_version": str(field(s.ol, "Version"))}
	if accs, err := getObj(s.ns, "Accounts"); err == nil {
		o["accounts"] = num(field(accs, "Count"))
		if a, err := callObj(accs, "Item", 1); err == nil {
			o["email"] = str(field(a, "SmtpAddress"))
			if cu, err := getObj(a, "CurrentUser"); err == nil {
				o["name"] = str(field(cu, "Name"))
			}
		}
	} else if isTransient(err) {
		return err
	}
	if str(o["name"]) == "" {
		if cu, err := getObj(s.ns, "CurrentUser"); err == nil {
			o["name"] = str(field(cu, "Name"))
			if str(o["email"]) == "" {
				if ae, err := getObj(cu, "AddressEntry"); err == nil {
					o["email"] = smtpOf(ae)
				}
			}
		}
	}
	if st, err := getObj(s.ns, "DefaultStore"); err == nil {
		o["store"] = str(field(st, "DisplayName"))
	}
	if v, ok := field(s.ns, "ExchangeConnectionMode").(int); ok {
		if m, ok := connectionModes[v]; ok {
			o["connection"] = m
		} else {
			o["connection"] = fmt.Sprintf("mode %d", v)
		}
	}
	s.emit(o)
	return nil
}
