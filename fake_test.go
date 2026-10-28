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
		sort.SliceStable(list, func(i, j int) bool {
			return list[i].props["ReceivedTime"].(time.Time).After(list[j].props["ReceivedTime"].(time.Time))
		})
		return nil, nil
	})
	it.on("Restrict", func(args ...any) (any, error) {
		m := regexp.MustCompile(`LIKE '%(.*?)%'`).FindStringSubmatch(args[0].(string))
		var hit []*fake
		for _, x := range list {
			if strings.Contains(x.props["Subject"].(string), m[1]) {
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
}

func newMailbox() *mailbox {
	mb := &mailbox{byID: map[string]*fake{}, dflt: map[int]*fake{}}
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

	mb.ol = newFake(nil).on("CreateItem", func(...any) (any, error) {
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
	})
	m.props["Attachments"] = fakeCol(&atts)
	m.props["_atts"] = &atts
	m.props["Attachments"].(*fake).on("Add", func(args ...any) (any, error) {
		a := newFake(map[string]any{"FileName": args[0].(string)})
		atts = append(atts, a)
		return a, nil
	})
	m.props["Recipients"] = newFake(nil).on("ResolveAll", func(...any) (any, error) { return true, nil })
	m.on("Save", func(...any) (any, error) {
		if m.props["EntryID"] == nil {
			m.props["EntryID"] = mb.id("NEW")
			mb.byID[m.props["EntryID"].(string)] = m
		}
		return nil, nil
	})
	reply := func(...any) (any, error) {
		r := mb.mail(nil, "AW: "+subject, time.Time{})
		r.props["BodyFormat"] = olFormatHTML
		r.props["HTMLBody"] = `<html><body><div class=WordSection1><p>Zitat</p></div></body></html>`
		return r, nil
	}
	m.on("Reply", reply).on("ReplyAll", reply)
	m.on("Delete", func(...any) (any, error) { delete(mb.byID, m.props["EntryID"].(string)); return nil, nil })
	m.on("Move", func(args ...any) (any, error) {
		delete(mb.byID, m.props["EntryID"].(string))
		m.props["EntryID"] = mb.id("M")
		mb.byID[m.props["EntryID"].(string)] = m
		return m, nil
	})
	if folder != nil {
		m.props["EntryID"] = mb.id("M")
		mb.byID[m.props["EntryID"].(string)] = m
		ms := folder.props["_mails"].(*[]*fake)
		*ms = append(*ms, m)
	}
	return m
}

func (mb *mailbox) session() *session { return &session{ol: mb.ol, ns: mb.ns} }

type fakeBackend struct {
	mb     *mailbox
	closed *int
}

func (b fakeBackend) begin() (obj, obj) { return b.mb.ol, b.mb.ns }
func (b fakeBackend) end()              {}
func (b fakeBackend) close()            { *b.closed++ }
