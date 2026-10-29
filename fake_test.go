package main

import (
	"fmt"
	"regexp"
	"sort"
	"strings"
	"time"
)

// Faelschung von Outlook hinter obj: gerade so viel Objektmodell, wie die Ops
// anfassen. Ein Property kann ein fester Wert oder eine func() any sein
// (dynamisch, z. B. Items), eine Methode eine func(args ...any) (any, error).

type fake struct {
	props map[string]any
	meths map[string]func(args ...any) (any, error)
	puts  []string // Protokoll der gesetzten Properties
}

func newFake(props map[string]any) *fake {
	if props == nil {
		props = map[string]any{}
	}
	return &fake{props: props, meths: map[string]func(args ...any) (any, error){}}
}

func (f *fake) get(name string, args ...any) (any, error) {
	if v, ok := f.props[name]; ok {
		if fn, ok := v.(func() any); ok {
			return fn(), nil
		}
		return v, nil
	}
	if fn, ok := f.meths[name]; ok {
		return fn(args...)
	}
	return nil, &comError{0x80020006, "Unbekannter Name: " + name}
}

func (f *fake) call(name string, args ...any) (any, error) {
	fn, ok := f.meths[name]
	if !ok {
		return nil, &comError{0x80020006, "Unbekannte Methode: " + name}
	}
	return fn(args...)
}

func (f *fake) put(name string, v any) error {
	f.props[name] = v
	f.puts = append(f.puts, name)
	return nil
}

func (f *fake) wasPut(name string) bool {
	for _, p := range f.puts {
		if p == name {
			return true
		}
	}
	return false
}

func (f *fake) on(name string, fn func(args ...any) (any, error)) *fake {
	f.meths[name] = fn
	return f
}

// fakeCol ist eine 1-basierte Collection ueber einen Slice, der wachsen darf.
func fakeCol(list *[]*fake) *fake {
	c := newFake(map[string]any{"Count": func() any { return len(*list) }})
	c.on("Item", func(args ...any) (any, error) {
		i := args[0].(int)
		if i < 1 || i > len(*list) {
			return nil, &comError{0x80020009, "Index ausserhalb"}
		}
		return (*list)[i-1], nil
	})
	return c
}

// fakeItems: GetFirst/GetNext, Count, Sort und ein Restrict, das den
// Suchbegriff aus dem DASL-Filter zieht und im Betreff sucht.
func fakeItems(mails []*fake) *fake {
	list := append([]*fake(nil), mails...)
	pos := 0
	it := newFake(map[string]any{"Count": func() any { return len(list) }})
	next := func(args ...any) (any, error) {
		if pos >= len(list) {
			return nil, nil
		}
		pos++
		return list[pos-1], nil
	}
	it.on("GetFirst", func(args ...any) (any, error) { pos = 0; return next() })
	it.on("GetNext", next)
	it.on("Sort", func(args ...any) (any, error) {
		if args[0] == "[Subject]" {
			sort.SliceStable(list, func(i, j int) bool {
				return list[i].props["Subject"].(string) < list[j].props["Subject"].(string)
			})
			return nil, nil
		}
		sort.SliceStable(list, func(i, j int) bool {
			return list[i].props["ReceivedTime"].(time.Time).After(list[j].props["ReceivedTime"].(time.Time))
		})
		return nil, nil
	})
	it.on("Restrict", func(args ...any) (any, error) {
		m := regexp.MustCompile(`LIKE '%(.*?)%'`).FindStringSubmatch(args[0].(string))
		var hit []*fake
		for _, x := range list {
			if strings.Contains(strings.ToLower(x.props["Subject"].(string)), strings.ToLower(m[1])) {
				hit = append(hit, x)
			}
		}
		return fakeItems(hit), nil
	})
	return it
}

// mailbox baut ein Postfach mit Ordnerbaum, Mails und Namespace.
type mailbox struct {
	ol, ns *fake
	root   *fake
	byID   map[string]*fake
	dflt   map[int]*fake
	nextID int
	gal    []*fake          // Firmen-Adressbuch
	known  map[string]*fake // was CreateRecipient(...).Resolve() eindeutig findet
	scans  int              // Durchlaeufe ueber die GAL (GetFirst)
}

func newMailbox() *mailbox {
	mb := &mailbox{byID: map[string]*fake{}, dflt: map[int]*fake{}, known: map[string]*fake{}}
	mb.root = mb.folder(nil, "user@example.com", 0)
	store := newFake(nil).on("GetRootFolder", func(...any) (any, error) { return mb.root, nil })

	mb.ns = newFake(map[string]any{"DefaultStore": store})
	mb.ns.on("GetDefaultFolder", func(args ...any) (any, error) {
		if f, ok := mb.dflt[args[0].(int)]; ok {
			return f, nil
		}
		return nil, &comError{0x80004005, "Standardordner fehlt"}
	})
	mb.ns.on("GetFolderFromID", func(args ...any) (any, error) {
		if f, ok := mb.byID[args[0].(string)]; ok && f.props["Folders"] != nil {
			return f, nil
		}
		return nil, &comError{0x80040107, "Ungültige ID."}
	})
	mb.ns.on("GetItemFromID", func(args ...any) (any, error) {
		if m, ok := mb.byID[args[0].(string)]; ok && m.props["Folders"] == nil {
			return m, nil
		}
		return nil, &comError{0x8004010F, "Die Nachricht wurde nicht gefunden."}
	})

	mb.ns.on("GetAddressEntryFromID", func(args ...any) (any, error) {
		for _, e := range mb.gal {
			if e.props["ID"] == args[0] {
				return e, nil
			}
		}
		return nil, &comError{0x80040107, "Ungültige ID."}
	})
	mb.ns.on("GetGlobalAddressList", func(...any) (any, error) {
		ents := fakeItems(mb.gal)
		first := ents.meths["GetFirst"]
		ents.on("GetFirst", func(args ...any) (any, error) { mb.scans++; return first(args...) })
		return newFake(map[string]any{"Name": "Global Address List", "AddressEntries": ents}), nil
	})
	mb.ns.on("CreateRecipient", func(args ...any) (any, error) {
		name := args[0].(string)
		r := newFake(map[string]any{"Name": name, "Address": ""})
		r.on("Resolve", func(...any) (any, error) {
			ae, ok := mb.known[strings.ToLower(name)]
			if ok {
				r.props["AddressEntry"] = ae
				r.props["Name"] = ae.props["Name"]
			}
			return ok, nil
		})
		return r, nil
	})
	// Namespace.CurrentUser bricht wie am echten Outlook kurz nach dem Start ab.
	mb.ns.on("CurrentUser", func(...any) (any, error) { return nil, &comError{0x80004004, "Vorgang abgebrochen"} })
	me := newFake(map[string]any{"Name": "Test User"})
	accs := []*fake{newFake(map[string]any{"SmtpAddress": "user@example.com", "CurrentUser": me})}
	mb.ns.props["Accounts"] = fakeCol(&accs)
	mb.ns.props["ExchangeConnectionMode"] = 700

	mb.ol = newFake(map[string]any{"Version": "16.0.0.1"}).on("CreateItem", func(...any) (any, error) {
		return mb.mail(nil, "", time.Time{}), nil
	})
	return mb
}

func (mb *mailbox) id(prefix string) string {
	mb.nextID++
	return fmt.Sprintf("%s%03d", prefix, mb.nextID)
}

// folder legt einen Ordner an; kind ist DefaultItemType (0 = Mail).
func (mb *mailbox) folder(parent *fake, name string, kind int) *fake {
	var subs []*fake
	var mails []*fake
	f := newFake(map[string]any{
		"Name": name, "EntryID": mb.id("F"), "DefaultItemType": kind,
		"UnReadItemCount": 0, "FolderPath": `\\` + name,
	})
	f.props["Folders"] = fakeCol(&subs)
	f.props["Items"] = func() any { return fakeItems(mails) }
	f.props["_mails"] = &mails
	f.props["_subs"] = &subs
	f.props["Folders"].(*fake).on("Add", func(args ...any) (any, error) {
		return mb.folder(f, args[0].(string), 0), nil
	})
	if parent != nil {
		ps := parent.props["_subs"].(*[]*fake)
		*ps = append(*ps, f)
		f.props["FolderPath"] = parent.props["FolderPath"].(string) + `\` + name
	}
	mb.byID[f.props["EntryID"].(string)] = f
	return f
}

// mail legt eine Mail an; ohne Ordner ist es ein neuer, ungespeicherter Entwurf.
func (mb *mailbox) mail(folder *fake, subject string, received time.Time) *fake {
	var atts []*fake
	m := newFake(map[string]any{
		"Subject": subject, "SenderName": "Absender", "To": "user@example.com",
		"ReceivedTime": received, "Body": "Text von " + subject, "UnRead": true,
		"BodyFormat": 1, "Categories": "",
		"SenderEmailType": "SMTP", "SenderEmailAddress": "sender@example.com",
	})
	m.props["Attachments"] = fakeCol(&atts)
	m.props["_atts"] = &atts
	m.props["Attachments"].(*fake).on("Add", func(args ...any) (any, error) {
		a := newFake(map[string]any{"FileName": args[0].(string)})
		atts = append(atts, a)
		return a, nil
	})
	m.props["Class"] = olMail
	m.props["Recipients"] = func() any { return fakeRecipients(m) }
	m.on("Save", func(...any) (any, error) {
		if m.props["EntryID"] == nil {
			m.props["EntryID"] = mb.id("NEW")
			mb.byID[m.props["EntryID"].(string)] = m
			// Neue Elemente speichert Outlook in "Entwuerfe".
			if d, ok := mb.dflt[olFolderDrafts]; ok {
				m.props["Parent"] = d
			}
		}
		return nil, nil
	})
	reply := func(prefix string) func(...any) (any, error) {
		return func(...any) (any, error) {
			r := mb.mail(nil, prefix+subject, time.Time{})
			r.props["BodyFormat"] = olFormatHTML
			r.props["HTMLBody"] = `<html><body><div class=WordSection1><p>Zitat</p></div></body></html>`
			return r, nil
		}
	}
	m.on("Reply", reply("AW: ")).on("ReplyAll", reply("AW: ")).on("Forward", reply("WG: "))
	// Send: wie echt wandert das Element in den Postausgang, die EntryID gilt
	// nicht mehr.
	m.on("Send", func(...any) (any, error) {
		m.props["_sent"] = true
		delete(mb.byID, str(m.props["EntryID"]))
		return nil, nil
	})
	m.on("Delete", func(...any) (any, error) { delete(mb.byID, m.props["EntryID"].(string)); return nil, nil })
	m.on("Move", func(args ...any) (any, error) {
		delete(mb.byID, m.props["EntryID"].(string))
		m.props["EntryID"] = mb.id("M")
		mb.byID[m.props["EntryID"].(string)] = m
		return m, nil
	})
	if folder != nil {
		m.props["Parent"] = folder
		m.props["EntryID"] = mb.id("M")
		mb.byID[m.props["EntryID"].(string)] = m
		ms := folder.props["_mails"].(*[]*fake)
		*ms = append(*ms, m)
	}
	return m
}

// fakeRecipients baut die Empfaenger aus To/CC/BCC. Aufgeloest ist, was ein
// "@" enthaelt.
func fakeRecipients(m *fake) *fake {
	var list []*fake
	for _, f := range []struct {
		prop string
		kind int
	}{{"To", 1}, {"CC", olCC}, {"BCC", olBCC}} {
		for _, a := range strings.Split(str(m.props[f.prop]), ";") {
			if a = strings.TrimSpace(a); a != "" {
				list = append(list, newFake(map[string]any{
					"Name": a, "Address": a, "Type": f.kind, "Resolved": strings.Contains(a, "@"),
				}))
			}
		}
	}
	rs := fakeCol(&list)
	rs.on("ResolveAll", func(...any) (any, error) {
		for _, r := range list {
			if r.props["Resolved"] != true {
				return false, nil
			}
		}
		return true, nil
	})
	return rs
}

// person legt einen Exchange-Benutzer im Firmen-Adressbuch an. Address ist
// wie echt nur der X.500-Name; die SMTP-Adresse gibt es ueber GetExchangeUser.
func (mb *mailbox) person(name, smtp, dept string) *fake {
	xu := newFake(map[string]any{
		"PrimarySmtpAddress": smtp, "Department": dept, "CompanyName": "Example Ltd",
		"BusinessTelephoneNumber": "+41 00 000 00 00",
		"StreetAddress":           "Street 1", "PostalCode": "8000", "City": "Zurich",
	})
	e := newFake(map[string]any{
		"Name": name, "ID": mb.id("GAL"), "AddressEntryUserType": olExchangeUserAddressEntry,
		"Type": "EX", "Address": "/O=EXAMPLE/OU=EXCHANGE/CN=RECIPIENTS/CN=" + strings.ToUpper(smtp),
	})
	e.on("GetExchangeUser", func(...any) (any, error) { return xu, nil })
	mb.gal = append(mb.gal, e)
	return e
}

// contact legt einen eigenen Kontakt an (Subject = FullName wie in Outlook).
func (mb *mailbox) contact(folder *fake, name, email, company string) *fake {
	c := newFake(map[string]any{
		"EntryID": mb.id("C"), "Class": olContact, "Subject": name, "FullName": name, "FileAs": name,
		"CompanyName": company, "Email1Address": email, "Email1AddressType": "SMTP",
		"Body": "Notiz zu " + name, "BusinessAddress": "Weg 2\r\n8000 Zurich",
	})
	mb.byID[c.props["EntryID"].(string)] = c
	ms := folder.props["_mails"].(*[]*fake)
	*ms = append(*ms, c)
	return c
}

func (mb *mailbox) session() *session { return &session{ol: mb.ol, ns: mb.ns, dir: &dirCache{}} }

type fakeBackend struct {
	mb     *mailbox
	closed *int
}

func (b fakeBackend) begin() (obj, obj) { return b.mb.ol, b.mb.ns }
func (b fakeBackend) end()              {}
func (b fakeBackend) close()            { *b.closed++ }
