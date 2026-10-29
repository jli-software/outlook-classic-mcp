package main

import (
	"regexp"
	"sort"
	"strings"
	"testing"
	"time"
)

func tm(day, hour, min int) time.Time { return time.Date(2026, 11, day, hour, min, 0, 0, time.UTC) }

// calBox: Postfach mit Kalender. Termine haengen in cal; eine Serie (every > 0)
// liefert mit IncludeRecurrences ihre Vorkommen, alle mit der EntryID der Serie.
type calBox struct {
	*testBox
	cal     []*fake
	created []*fake // ueber CreateItem(olAppointmentItem) angelegt
	steps   []string
}

func newCalBox(t *testing.T) *calBox {
	old := jetTime
	jetTime = func(x time.Time) string { return x.Format("2006-01-02 15:04") }
	t.Cleanup(func() { jetTime = old })
	b := &calBox{testBox: newTestBox()}
	b.dflt[olFolderCalendar] = b.calendar
	b.calendar.props["Items"] = func() any { return b.calItems() }
	create := b.ol.meths["CreateItem"]
	b.ol.on("CreateItem", func(args ...any) (any, error) {
		if args[0] == olAppointmentItem {
			a := b.newAppt("", time.Time{}, time.Time{})
			b.created = append(b.created, a)
			return a, nil
		}
		return create(args...)
	})
	return b
}

// appt legt einen gespeicherten Termin an.
func (b *calBox) appt(subject string, start, end time.Time) *fake {
	a := b.newAppt(subject, start, end)
	a.call("Save")
	return a
}

// newAppt: ungespeichert, ohne EntryID — wie frisch aus CreateItem.
func (b *calBox) newAppt(subject string, start, end time.Time) *fake {
	var recips []*fake
	a := newFake(map[string]any{
		"EntryID": nil, "Class": olAppointment, "Subject": subject, "Start": start, "End": end,
		"Location": "", "AllDayEvent": false, "BusyStatus": 2, "MeetingStatus": olNonMeeting,
		"ResponseStatus": 0, "Organizer": "Test User", "IsRecurring": false, "RecurrenceState": 0,
		"Body": "", "Categories": "", "ReminderSet": true, "ReminderMinutesBeforeStart": 15,
	})
	rs := fakeCol(&recips)
	rs.on("Add", func(args ...any) (any, error) {
		addr := args[0].(string)
		r := newFake(map[string]any{"Name": addr, "Type": olRequired, "Resolved": strings.Contains(addr, "@"), "MeetingResponseStatus": 0})
		r.props["AddressEntry"] = newFake(map[string]any{"Type": "SMTP", "Address": addr, "AddressEntryUserType": 30})
		recips = append(recips, r)
		return r, nil
	})
	rs.on("ResolveAll", func(...any) (any, error) { return true, nil })
	a.props["Recipients"] = rs
	a.on("Save", func(...any) (any, error) {
		if a.props["EntryID"] == nil {
			a.props["EntryID"] = b.id("A")
			b.byID[a.props["EntryID"].(string)] = a
			b.cal = append(b.cal, a)
		}
		return nil, nil
	})
	a.on("Delete", func(...any) (any, error) {
		a.props["_deleted"] = true
		return nil, nil
	})
	a.on("Respond", func(args ...any) (any, error) {
		r := b.mail(nil, "Accepted: "+subject, time.Time{})
		r.props["_response"] = args[0]
		r.on("Send", func(...any) (any, error) { panic("Respond darf nie senden") })
		return r, nil
	})
	return a
}

// series macht aus a eine taegliche Serie mit n Vorkommen.
func (b *calBox) series(a *fake, n int) []*fake {
	a.props["IsRecurring"] = true
	a.props["RecurrenceState"] = olApptMaster
	var occ []*fake
	for i := 0; i < n; i++ {
		o := newFake(map[string]any{})
		for k, v := range a.props {
			o.props[k] = v
		}
		for k, v := range a.meths {
			o.meths[k] = v
		}
		o.props["RecurrenceState"] = olApptOccurrence
		o.props["Start"] = a.props["Start"].(time.Time).AddDate(0, 0, i)
		o.props["End"] = a.props["End"].(time.Time).AddDate(0, 0, i)
		o.on("Delete", func(...any) (any, error) { o.props["_deleted"] = true; return nil, nil })
		occ = append(occ, o)
	}
	a.props["_occ"] = occ
	pat := newFake(nil).on("GetOccurrence", func(args ...any) (any, error) {
		for _, o := range occ {
			if o.props["Start"].(time.Time).Equal(args[0].(time.Time)) {
				return o, nil
			}
		}
		return nil, &comError{0x80004005, "Kein Vorkommen"}
	})
	var none []*fake
	pat.props["Exceptions"] = fakeCol(&none)
	a.on("GetRecurrencePattern", func(...any) (any, error) { return pat, nil })
	return occ
}

// calItems verlangt wie Outlook Sort("[Start]") und IncludeRecurrences vor
// Restrict — sonst gibt es nur die Serie selbst, ohne Vorkommen.
func (b *calBox) calItems() obj {
	sorted, recur := false, false
	it := fakeItems(nil)
	it.on("Sort", func(args ...any) (any, error) {
		b.steps = append(b.steps, "Sort "+args[0].(string))
		sorted = args[0] == "[Start]" && !recur
		return nil, nil
	})
	it.on("Restrict", func(args ...any) (any, error) {
		b.steps = append(b.steps, "Restrict")
		m := regexp.MustCompile(`^\[Start\] < '([^']+)' AND \[End\] > '([^']+)'$`).FindStringSubmatch(args[0].(string))
		if m == nil {
			panic("unerwarteter Filter: " + args[0].(string))
		}
		to, _ := time.Parse("2006-01-02 15:04", m[1])
		from, _ := time.Parse("2006-01-02 15:04", m[2])
		recur = recur && sorted
		var hit []*fake
		for _, a := range b.cal {
			list := []*fake{a}
			if occ, ok := a.props["_occ"].([]*fake); ok && recur {
				list = occ
			}
			for _, x := range list {
				if x.props["_deleted"] == nil && x.props["Start"].(time.Time).Before(to) && x.props["End"].(time.Time).After(from) {
					hit = append(hit, x)
				}
			}
		}
		sort.SliceStable(hit, func(i, j int) bool {
			return hit[i].props["Start"].(time.Time).Before(hit[j].props["Start"].(time.Time))
		})
		res := fakeItems(hit)
		if recur {
			next := res.meths["GetNext"]
			res.on("GetNext", func(args ...any) (any, error) {
				v, err := next(args...)
				if v == nil && err == nil {
					return nil, &comError{1, "Unzulässige Funktion."}
				}
				return v, err
			})
		}
		return res, nil
	})
	return &fakePut{it, func(name string, v any) {
		if name == "IncludeRecurrences" {
			b.steps = append(b.steps, "IncludeRecurrences")
			recur = v == true
		}
	}}
}

// fakePut meldet gesetzte Properties an einen Beobachter.
type fakePut struct {
	*fake
	onPut func(string, any)
}

func (f *fakePut) put(name string, v any) error { f.onPut(name, v); return f.fake.put(name, v) }

func calRun(t *testing.T, s *session, op string, p map[string]any) []calendarItem {
	t.Helper()
	if err := s.do(op, p); err != nil {
		t.Fatalf("%s: %v", op, err)
	}
	return parseJSONLines[calendarItem](s.out.Bytes())
}

func TestCalendarListIncludesOccurrences(t *testing.T) {
	b := newCalBox(t)
	b.appt("Vorher", tm(1, 9, 0), tm(1, 10, 0))
	b.appt("Planung", tm(3, 14, 0), tm(3, 15, 0))
	daily := b.appt("Daily", tm(2, 8, 30), tm(2, 8, 45))
	b.series(daily, 5)

	got := calRun(t, b.session(), "calendar_list", map[string]any{"start": "2026-11-02", "end": "2026-11-04"})
	var subj []string
	for _, c := range got {
		subj = append(subj, c.Subject)
	}
	if strings.Join(subj, ",") != "Daily,Daily,Planung,Daily" {
		t.Fatalf("Termine: %v (Schritte %v)", subj, b.steps)
	}
	if strings.Join(b.steps, ",") != "Sort [Start],IncludeRecurrences,Restrict" {
		t.Errorf("Reihenfolge: %v", b.steps)
	}
	entry := daily.props["EntryID"].(string)
	if got[0].ID != entry+"|2026-11-02T08:30" || got[1].ID != entry+"|2026-11-03T08:30" || !got[0].Recurring {
		t.Errorf("Vorkommen-IDs: %q %q", got[0].ID, got[1].ID)
	}
	if !strings.HasPrefix(got[0].Start, "2026-11-02T08:30:00") || got[0].ShowAs != "busy" {
		t.Errorf("Start/ShowAs: %+v", got[0])
	}

	// top und query
	got = calRun(t, b.session(), "calendar_list", map[string]any{"start": "2026-11-01", "end": "2026-11-10", "top": 2})
	if len(got) != 2 || got[0].Subject != "Vorher" {
		t.Errorf("top: %+v", got)
	}
	got = calRun(t, b.session(), "calendar_list", map[string]any{"start": "2026-11-01", "end": "2026-11-10", "query": "plan"})
	if len(got) != 1 || got[0].Subject != "Planung" {
		t.Errorf("query: %+v", got)
	}
	if err := b.session().do("calendar_list", map[string]any{"start": "morgen"}); err == nil {
		t.Error("ungueltige Zeit ohne Fehler")
	}
}

func TestCalendarGetOccurrenceAndDetails(t *testing.T) {
	b := newCalBox(t)
	m := b.appt("Review", tm(5, 10, 0), tm(5, 11, 0))
	m.props["MeetingStatus"] = olMeetingReceived
	m.props["ResponseStatus"] = 3
	m.props["Organizer"] = "Beat Muster"
	m.props["Body"] = "Agenda\r\nJoin: https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0?context=x\r\n"
	rs := m.props["Recipients"].(*fake)
	rs.call("Add", "beat@example.com")
	opt, _ := rs.call("Add", "anna@example.com")
	opt.(*fake).props["Type"] = olOptional
	opt.(*fake).props["MeetingResponseStatus"] = 2
	daily := b.appt("Daily", tm(2, 8, 30), tm(2, 8, 45))
	b.series(daily, 3)
	entry := daily.props["EntryID"].(string)

	got := calRun(t, b.session(), "calendar_get", map[string]any{"ids": []string{
		m.props["EntryID"].(string), entry + "|2026-11-03T08:30", entry + "|2026-11-09T08:30", "BAD"}})
	if len(got) != 4 {
		t.Fatalf("got %d: %+v", len(got), got)
	}
	r := got[0]
	if r.Meeting != "attendee" || r.Response != "accepted" || r.Organizer != "Beat Muster" ||
		!strings.HasPrefix(r.TeamsURL, "https://teams.microsoft.com/l/meetup-join/") || strings.ContainsAny(r.TeamsURL, "\r\n") {
		t.Errorf("Details: %+v", r)
	}
	if len(r.Attendees) != 2 || r.Attendees[1].Type != "optional" || r.Attendees[1].Response != "tentative" || r.Attendees[1].Email != "anna@example.com" {
		t.Errorf("Teilnehmer: %+v", r.Attendees)
	}
	if r.Reminder == nil || *r.Reminder != 15 {
		t.Errorf("Erinnerung: %v", r.Reminder)
	}
	if !strings.HasPrefix(got[1].Start, "2026-11-03T08:30") {
		t.Errorf("Vorkommen: %+v", got[1])
	}
	if got[2].Error == "" || got[3].Error == "" {
		t.Errorf("Fehler je ID: %+v / %+v", got[2], got[3])
	}
}

func TestCalendarCreate(t *testing.T) {
	b := newCalBox(t)
	got := calRun(t, b.session(), "calendar_create", map[string]any{
		"subject": "Zahnarzt", "start": "2026-11-04T07:30", "location": "Praxis", "reminder": float64(-1), "show_as": "oof",
	})
	a := b.created[0]
	if len(got) != 1 || got[0].NotSent || a.props["End"] != tm(4, 8, 0) || a.props["Location"] != "Praxis" ||
		a.props["ReminderSet"] != false || a.props["BusyStatus"] != 3 || a.wasPut("MeetingStatus") {
		t.Errorf("Termin: %+v / %v", got, a.props)
	}

	// Ganztaegig ueber zwei Tage: Enddatum zaehlt einschliesslich.
	calRun(t, b.session(), "calendar_create", map[string]any{"subject": "Ferien", "start": "2026-11-09", "end": "2026-11-10"})
	a = b.created[1]
	if a.props["AllDayEvent"] != true || a.props["Start"] != tm(9, 0, 0) || a.props["End"] != tm(11, 0, 0) {
		t.Errorf("ganztaegig: %v – %v", a.props["Start"], a.props["End"])
	}

	// Mit Teilnehmern: Besprechung, gespeichert, nicht gesendet.
	got = calRun(t, b.session(), "calendar_create", map[string]any{
		"subject": "Sync", "start": "2026-11-05T13:00+00:00", "duration": float64(45),
		"attendees": []any{"beat@example.com", "Nico"}, "optional_attendees": []any{"anna@example.com"},
	})
	a = b.created[2]
	if a.props["MeetingStatus"] != olMeeting || !got[0].NotSent || got[0].Unresolved != "Nico" {
		t.Errorf("Besprechung: %+v", got)
	}
	start := a.props["Start"].(time.Time)
	want := wall(time.Date(2026, 11, 5, 13, 0, 0, 0, time.UTC).In(time.Local))
	if !start.Equal(want) || !a.props["End"].(time.Time).Equal(want.Add(45*time.Minute)) {
		t.Errorf("Zone nicht nach Ortszeit: %v, want %v", start, want)
	}
	recips := a.props["Recipients"].(*fake)
	if n, _ := getInt(recips, "Count"); n != 3 {
		t.Errorf("%d Empfaenger", n)
	}
	if r, _ := callObj(recips, "Item", 3); num(field(r, "Type")) != olOptional {
		t.Error("optionaler Teilnehmer nicht optional")
	}

	for _, p := range []map[string]any{
		{"subject": "x"},
		{"subject": "x", "start": "2026-11-04T10:00", "end": "2026-11-04T09:00"},
		{"subject": "x", "start": "2026-11-04T10:00", "show_as": "maybe"},
	} {
		if err := b.session().do("calendar_create", p); err == nil {
			t.Errorf("%v: kein Fehler", p)
		}
	}
}

func TestCalendarUpdate(t *testing.T) {
	b := newCalBox(t)
	a := b.appt("Alt", tm(6, 9, 0), tm(6, 10, 30))
	id := a.props["EntryID"].(string)
	calRun(t, b.session(), "calendar_update", map[string]any{"id": id, "start": "2026-11-07T14:00", "subject": "Neu"})
	if a.props["Start"] != tm(7, 14, 0) || a.props["End"] != tm(7, 15, 30) || a.props["Subject"] != "Neu" {
		t.Errorf("verschoben: %v – %v %v", a.props["Start"], a.props["End"], a.props["Subject"])
	}
	if a.wasPut("Body") || a.wasPut("Location") || a.wasPut("AllDayEvent") {
		t.Errorf("nicht uebergebene Felder angefasst: %v", a.puts)
	}

	// Verschickte eigene Besprechung: Hinweis, dass niemand benachrichtigt wird.
	a.props["MeetingStatus"] = olMeeting
	a.props["PropertyAccessor"] = newFake(nil).on("GetProperty", func(...any) (any, error) { return true, nil })
	got := calRun(t, b.session(), "calendar_update", map[string]any{"id": id, "location": "Raum 2"})
	if !strings.Contains(got[0].Note, "not notified") {
		t.Errorf("Hinweis fehlt: %+v", got)
	}

	// Ein Vorkommen einer Serie.
	d := b.appt("Daily", tm(2, 8, 30), tm(2, 8, 45))
	occ := b.series(d, 3)
	calRun(t, b.session(), "calendar_update", map[string]any{"id": d.props["EntryID"].(string) + "|2026-11-03T08:30", "duration": float64(30)})
	if occ[1].props["End"] != tm(3, 9, 0) || occ[0].props["End"] != tm(2, 8, 45) || d.wasPut("End") {
		t.Error("falsches Vorkommen geaendert")
	}
}

func TestCalendarDeleteGuardsSentMeetings(t *testing.T) {
	b := newCalBox(t)
	plain := b.appt("Allein", tm(6, 9, 0), tm(6, 10, 0))
	own := b.appt("Meine Besprechung", tm(6, 11, 0), tm(6, 12, 0))
	own.props["MeetingStatus"] = olMeeting // kein PropertyAccessor: zaehlt als verschickt
	draft := b.appt("Ungesendet", tm(6, 13, 0), tm(6, 14, 0))
	draft.props["MeetingStatus"] = olMeeting
	draft.props["PropertyAccessor"] = newFake(nil).on("GetProperty", func(...any) (any, error) { return false, nil })
	recv := b.appt("Eingeladen", tm(6, 15, 0), tm(6, 16, 0))
	recv.props["MeetingStatus"] = olMeetingReceived

	id := func(f *fake) string { return f.props["EntryID"].(string) }
	got := calRun(t, b.session(), "calendar_delete", map[string]any{"ids": []string{id(plain), id(own), id(draft), id(recv)}})
	if len(got) != 4 || !got[0].Deleted || got[1].Error == "" || !got[2].Deleted || !got[3].Deleted || got[3].Note == "" {
		t.Fatalf("delete: %+v", got)
	}
	if own.props["_deleted"] != nil || plain.props["_deleted"] != true {
		t.Error("verschickte Besprechung geloescht oder Termin nicht")
	}
}

func TestCalendarRespondSavesDraft(t *testing.T) {
	b := newCalBox(t)
	inv := b.appt("Einladung", tm(8, 9, 0), tm(8, 10, 0))
	own := b.appt("Eigener", tm(8, 11, 0), tm(8, 12, 0))
	inv.props["MeetingStatus"] = olMeetingReceived

	got := calRun(t, b.session(), "calendar_respond", map[string]any{"id": inv.props["EntryID"].(string), "response": "Decline", "comment": "Bin weg"})
	if len(got) != 1 || !got[0].NotSent || got[0].Response != "declined" || got[0].ID == "" {
		t.Fatalf("respond: %+v", got)
	}
	r := b.byID[got[0].ID]
	if r.props["_response"] != 4 || r.props["Body"] != "Bin weg" {
		t.Errorf("Antwort: %v", r.props)
	}
	for _, p := range []map[string]any{
		{"id": own.props["EntryID"].(string), "response": "accept"},
		{"id": inv.props["EntryID"].(string), "response": "vielleicht"},
	} {
		if err := b.session().do("calendar_respond", p); err == nil {
			t.Errorf("%v: kein Fehler", p)
		}
	}
}

func TestCalendarFreeBusy(t *testing.T) {
	b := newCalBox(t)
	create := b.ns.meths["CreateRecipient"]
	b.ns.on("CreateRecipient", func(args ...any) (any, error) {
		r, _ := create(args...)
		f := r.(*fake)
		f.on("FreeBusy", func(a ...any) (any, error) {
			if args[0] == "ext@example.org" {
				return nil, &comError{0x80004005, "Keine Daten"}
			}
			if a[0] != tm(10, 0, 0) || a[1] != 30 {
				t.Errorf("FreeBusy(%v, %v)", a[0], a[1])
			}
			// 08:00–09:30 gebucht, 13:00–13:30 mit Vorbehalt
			return strings.Repeat("0", 16) + "222" + strings.Repeat("0", 7) + "1" + strings.Repeat("0", 200), nil
		})
		return f, nil
	})
	out, err := func() ([]freeBusy, error) {
		s := b.session()
		err := s.do("calendar_free_busy", map[string]any{"emails": []any{"beat@example.com", "ext@example.org"}, "date": "2026-11-10"})
		return parseJSONLines[freeBusy](s.out.Bytes()), err
	}()
	if err != nil || len(out) != 2 {
		t.Fatalf("%v %+v", err, out)
	}
	bs := out[0].Busy
	if len(bs) != 2 || !strings.HasPrefix(bs[0].Start, "2026-11-10T08:00") || !strings.HasPrefix(bs[0].End, "2026-11-10T09:30") ||
		bs[0].Status != "busy" || bs[1].Status != "tentative" {
		t.Errorf("Bloecke: %+v", bs)
	}
	if out[1].Error == "" {
		t.Errorf("Fehler fuer externe Adresse fehlt: %+v", out[1])
	}
}

func TestParseTime(t *testing.T) {
	for in, want := range map[string]time.Time{
		"2026-11-04T07:30":    tm(4, 7, 30),
		"2026-11-04 07:30":    tm(4, 7, 30),
		"2026-11-04T07:30:00": tm(4, 7, 30),
		"2026-11-04":          tm(4, 0, 0),
	} {
		got, _, err := parseTime(in)
		if err != nil || !got.Equal(want) {
			t.Errorf("parseTime(%q) = %v, %v", in, got, err)
		}
	}
	// Mit Zone: nach Ortszeit umgerechnet, als Wanduhr.
	got, _, _ := parseTime("2026-11-04T07:30:00Z")
	if want := wall(time.Date(2026, 11, 4, 7, 30, 0, 0, time.UTC).In(time.Local)); !got.Equal(want) {
		t.Errorf("Zone: %v, want %v", got, want)
	}
	if _, d, _ := parseTime("2026-11-04"); !d {
		t.Error("Datum ohne Uhrzeit nicht erkannt")
	}
	// Ausgabe: Wanduhr mit dem Offset der Ortszeit, Uhrzeit unveraendert.
	if s := isoOut(tm(4, 7, 30)); !strings.HasPrefix(s, "2026-11-04T07:30:00") {
		t.Errorf("isoOut: %s", s)
	}
}

// Datum fuer den Restrict-Filter: Gebietsschema, vierstelliges Jahr und
// immer mit Uhrzeit — auch um Mitternacht.
func TestLocalDate(t *testing.T) {
	for _, x := range []time.Time{tm(3, 0, 0), tm(3, 14, 30)} {
		s := localDate(x)
		day, clock, ok := strings.Cut(s, " ")
		if !ok || !strings.Contains(day, "2026") || !strings.Contains(clock, ":") {
			t.Errorf("localDate(%v) = %q", x, s)
		}
		if x.Hour() == 14 && !strings.Contains(clock, "30") {
			t.Errorf("Uhrzeit fehlt: %q", s)
		}
	}
}
