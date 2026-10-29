package main

import (
	"errors"
	"fmt"
	"regexp"
	"strings"
	"time"
)

// Kalender: Termine lesen, anlegen, aendern, loeschen, auf Einladungen
// antworten, Frei/Gebucht anderer abfragen. Gesendet wird nichts — eine
// Besprechung bleibt ungesendet, eine Antwort wird ein Entwurf (Falle 8).

const (
	olFolderCalendar  = 9
	olAppointmentItem = 1
	olAppointment     = 26 // OlObjectClass

	// OlMeetingStatus
	olNonMeeting                 = 0
	olMeeting                    = 1
	olMeetingReceived            = 3
	olMeetingCanceled            = 5
	olMeetingReceivedAndCanceled = 7

	// OlRecurrenceState
	olApptMaster     = 1
	olApptOccurrence = 2
	olApptException  = 3

	// OlMeetingRecipientType
	olRequired = 1
	olOptional = 2

	// FInvited: true, sobald die Einladungen verschickt sind.
	prFInvited = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8229000B"
)

var (
	showAsNames   = []string{"free", "tentative", "busy", "out of office", "working elsewhere"}
	responseNames = []string{"none", "organizer", "tentative", "accepted", "declined", "not responded"}
	responseCodes = map[string]int{"accept": 3, "accepted": 3, "tentative": 2, "decline": 4, "declined": 4}
	showAsCodes   = map[string]int{"free": 0, "tentative": 1, "busy": 2, "oof": 3, "out of office": 3, "away": 3, "working elsewhere": 4, "elsewhere": 4}
	teamsURL      = regexp.MustCompile(`https://teams\.(?:microsoft|live)\.com/(?:l/meetup-join|meet)/[^\s"<>]+`)
)

func nameOf(names []string, i int) string {
	if i >= 0 && i < len(names) {
		return names[i]
	}
	return fmt.Sprintf("status %d", i)
}

// ---- Zeiten ----
//
// Outlook rechnet in Ortszeit ohne Zone. Intern ist eine Zeit darum die
// Ortszeit als Wanduhr mit dem Etikett UTC — genau so, wie fromOADate sie
// liefert und wie go-ole sie zurueckschreibt (Format ohne Zone). Rein kommt
// ISO 8601; ohne Zone gilt Ortszeit, mit Zone wird nach Ortszeit umgerechnet.
// Raus geht ISO 8601 mit dem Offset der Ortszeit.

var timeLayouts = []string{
	time.RFC3339, "2006-01-02T15:04Z07:00",
	"2006-01-02T15:04:05", "2006-01-02T15:04", "2006-01-02 15:04:05", "2006-01-02 15:04",
}

// parseTime liefert die Wanduhr der Ortszeit; dateOnly bei "2026-10-29".
func parseTime(s string) (t time.Time, dateOnly bool, err error) {
	s = strings.TrimSpace(s)
	if d, err := time.Parse("2006-01-02", s); err == nil {
		return d, true, nil
	}
	for _, l := range timeLayouts {
		x, err := time.Parse(l, s)
		if err != nil {
			continue
		}
		if strings.Contains(l, "Z07") {
			x = x.In(time.Local)
		}
		return wall(x), false, nil
	}
	return time.Time{}, false, fmt.Errorf("invalid time %q — use ISO 8601, e.g. 2026-10-29T14:00 or 2026-10-29", s)
}

func wall(x time.Time) time.Time {
	return time.Date(x.Year(), x.Month(), x.Day(), x.Hour(), x.Minute(), x.Second(), 0, time.UTC)
}

// isoOut haengt der Wanduhr den Offset der Ortszeit an.
func isoOut(t time.Time) string {
	return time.Date(t.Year(), t.Month(), t.Day(), t.Hour(), t.Minute(), t.Second(), 0, time.Local).Format(time.RFC3339)
}

func today() time.Time {
	n := time.Now()
	return time.Date(n.Year(), n.Month(), n.Day(), 0, 0, 0, 0, time.UTC)
}

// jetTime: Datum fuer einen Restrict-Filter, im Gebietsschema des Benutzers
// (Falle 18). Eine Variable, damit die Tests ein festes Format setzen koennen.
var jetTime = localDate

// ---- IDs ----
//
// Mit IncludeRecurrences tragen alle Vorkommen einer Serie die EntryID der
// Serie. Ein Vorkommen bekommt darum die ID "EntryID|Start"; ohne Zusatz
// meint die ID die ganze Serie bzw. den Einzeltermin (Falle 18).

const occSep = "|"

func apptID(a obj) string {
	id := str(field(a, "EntryID"))
	if st := num(field(a, "RecurrenceState")); st == olApptOccurrence || st == olApptException {
		if t, ok := field(a, "Start").(time.Time); ok {
			id += occSep + t.Format("2006-01-02T15:04")
		}
	}
	return id
}

func (s *session) resolveAppt(id string) (obj, error) {
	entry, occ, isOcc := strings.Cut(id, occSep)
	a, err := callObj(s.ns, "GetItemFromID", entry)
	if err != nil {
		return nil, err
	}
	if num(field(a, "Class")) != olAppointment {
		return nil, fmt.Errorf("not a calendar item")
	}
	if !isOcc {
		return a, nil
	}
	t, _, err := parseTime(occ)
	if err != nil {
		return nil, err
	}
	pat, err := callObj(a, "GetRecurrencePattern")
	if err != nil {
		return nil, err
	}
	if o, err := callObj(pat, "GetOccurrence", t); err == nil {
		return o, nil
	} else if isTransient(err) {
		return nil, err
	}
	// Verschobene Vorkommen (Ausnahmen) findet GetOccurrence nur unter dem
	// urspruenglichen Datum; die ID traegt aber den aktuellen Beginn.
	var found obj
	if exs, err := getObj(pat, "Exceptions"); err == nil {
		err := each(exs, func(ex obj) (bool, error) {
			if field(ex, "Deleted") == true {
				return false, nil
			}
			if ai, err := getObj(ex, "AppointmentItem"); err == nil {
				if st, ok := field(ai, "Start").(time.Time); ok && st.Equal(t) {
					found = ai
					return true, nil
				}
			}
			return false, nil
		})
		if err != nil && isTransient(err) {
			return nil, err
		}
	}
	if found == nil {
		return nil, fmt.Errorf("occurrence not found (moved or deleted?) — list the calendar again")
	}
	return found, nil
}

// ---- Beschreiben ----

func apptInfo(a obj, full bool) map[string]any {
	o := map[string]any{
		"id":       apptID(a),
		"subject":  str(field(a, "Subject")),
		"location": str(field(a, "Location")),
		"all_day":  field(a, "AllDayEvent") == true,
		"show_as":  nameOf(showAsNames, num(field(a, "BusyStatus"))),
	}
	if t, ok := field(a, "Start").(time.Time); ok {
		o["start"] = isoOut(t)
	}
	if t, ok := field(a, "End").(time.Time); ok {
		o["end"] = isoOut(t)
	}
	switch num(field(a, "MeetingStatus")) {
	case olMeeting:
		o["meeting"] = "organizer"
	case olMeetingReceived:
		o["meeting"] = "attendee"
		o["response"] = nameOf(responseNames, num(field(a, "ResponseStatus")))
	case olMeetingCanceled, olMeetingReceivedAndCanceled:
		o["meeting"] = "canceled"
	}
	if o["meeting"] != nil {
		o["organizer"] = str(field(a, "Organizer"))
	}
	o["recurring"] = field(a, "IsRecurring") == true
	if !full {
		return o
	}
	body := str(field(a, "Body"))
	o["body"] = body
	o["categories"] = str(field(a, "Categories"))
	if field(a, "ReminderSet") == true {
		o["reminder"] = num(field(a, "ReminderMinutesBeforeStart"))
	} else {
		o["reminder"] = -1
	}
	if u := teamsURL.FindString(body + " " + str(field(a, "Location"))); u != "" {
		o["teams_url"] = u
	}
	var atts []map[string]any
	if rs, err := getObj(a, "Recipients"); err == nil {
		each(rs, func(r obj) (bool, error) {
			email := ""
			if ae, err := getObj(r, "AddressEntry"); err == nil {
				email = smtpOf(ae)
			}
			kind := "required"
			switch num(field(r, "Type")) {
			case olOptional:
				kind = "optional"
			case 3:
				kind = "resource"
			}
			atts = append(atts, map[string]any{
				"name": str(field(r, "Name")), "email": email, "type": kind,
				"response": nameOf(responseNames, num(field(r, "MeetingResponseStatus"))),
			})
			return false, nil
		})
	}
	if len(atts) > 0 {
		o["attendees"] = atts
	}
	return o
}

// ---- Lesen ----

// calendarList: Die Reihenfolge Sort → IncludeRecurrences → Restrict ist
// Pflicht. Ohne Sort nach [Start] liefert Outlook keine Serienvorkommen,
// Items.Count ist mit IncludeRecurrences bedeutungslos, das Datum im Filter
// muss im Gebietsschema stehen, und das Listenende kommt als Fehler (Falle 18).
func (s *session) calendarList(p map[string]any) error {
	start, end, err := listRange(pStr(p, "start"), pStr(p, "end"))
	if err != nil {
		return err
	}
	cal, err := callObj(s.ns, "GetDefaultFolder", olFolderCalendar)
	if err != nil {
		return err
	}
	all, err := getObj(cal, "Items")
	if err != nil {
		return err
	}
	if _, err := all.call("Sort", "[Start]"); err != nil {
		return err
	}
	if err := all.put("IncludeRecurrences", true); err != nil {
		return err
	}
	items, err := callObj(all, "Restrict", "[Start] < '"+jetTime(end)+"' AND [End] > '"+jetTime(start)+"'")
	if err != nil {
		return err
	}
	top, n := pInt(p, "top", 50), 0
	q := strings.Fields(strings.ToLower(pStr(p, "query")))
	err = eachItem(items, func(a obj) (bool, error) {
		if n >= top {
			return true, nil
		}
		// Outlook nimmt die Grenzen nicht genau (ein ganztaegiger Termin am
		// Folgetag rutscht mit) — darum hier nochmals pruefen.
		if st, ok := field(a, "Start").(time.Time); ok && !st.Before(end) {
			return false, nil
		}
		if en, ok := field(a, "End").(time.Time); ok && !en.After(start) {
			return false, nil
		}
		if len(q) > 0 && !matchWords(str(field(a, "Subject"))+" "+str(field(a, "Location")), q) {
			return false, nil
		}
		s.emit(apptInfo(a, false))
		n++
		return false, nil
	})
	// Mit IncludeRecurrences meldet GetFirst/GetNext das Ende der Liste als
	// Fehler 0x00000001 statt mit Nothing.
	var ce *comError
	if errors.As(err, &ce) && ce.code == 1 {
		return nil
	}
	return err
}

// listRange: Standard heute bis +7 Tage; ein Enddatum ohne Uhrzeit zaehlt
// bis Ende des Tages.
func listRange(from, to string) (time.Time, time.Time, error) {
	start := today()
	if from != "" {
		t, _, err := parseTime(from)
		if err != nil {
			return start, start, err
		}
		start = t
	}
	end := start.AddDate(0, 0, 7)
	if to != "" {
		t, dateOnly, err := parseTime(to)
		if err != nil {
			return start, end, err
		}
		if dateOnly {
			t = t.AddDate(0, 0, 1)
		}
		end = t
	}
	if !end.After(start) {
		return start, end, fmt.Errorf("end must be after start")
	}
	return start, end, nil
}

func (s *session) calendarGet(p map[string]any) error {
	return s.forEachID(p, func(id string) error {
		a, err := s.resolveAppt(id)
		if err != nil {
			return err
		}
		s.emit(apptInfo(a, true))
		return nil
	})
}

// ---- Schreiben ----

// setTimes setzt Beginn und Ende aus start/end/duration. Ein Datum ohne
// Uhrzeit macht einen ganztaegigen Termin; ein Enddatum zaehlt dann
// einschliesslich.
func setTimes(a obj, p map[string]any, creating bool) error {
	var start, end time.Time
	allDay := pBool(p, "all_day", false)
	if pHas(p, "start") {
		t, dateOnly, err := parseTime(pStr(p, "start"))
		if err != nil {
			return err
		}
		start, allDay = t, allDay || dateOnly
	} else if creating {
		return fmt.Errorf("start is required")
	} else if t, ok := field(a, "Start").(time.Time); ok {
		start = t
	}
	if allDay {
		start = time.Date(start.Year(), start.Month(), start.Day(), 0, 0, 0, 0, time.UTC)
	}
	switch {
	case pHas(p, "end"):
		t, dateOnly, err := parseTime(pStr(p, "end"))
		if err != nil {
			return err
		}
		if allDay && dateOnly {
			t = t.AddDate(0, 0, 1)
		}
		end = t
	case pInt(p, "duration", 0) > 0:
		end = start.Add(time.Duration(pInt(p, "duration", 0)) * time.Minute)
	case allDay:
		end = start.AddDate(0, 0, 1)
	case creating:
		end = start.Add(30 * time.Minute)
	default:
		// Verschieben: Dauer bleibt.
		if ostart, ok := field(a, "Start").(time.Time); ok {
			if oe, ok := field(a, "End").(time.Time); ok {
				end = start.Add(oe.Sub(ostart))
			}
		}
	}
	if !end.After(start) {
		return fmt.Errorf("end must be after start")
	}
	if creating || pHas(p, "all_day") || allDay {
		if err := a.put("AllDayEvent", allDay); err != nil {
			return err
		}
	}
	// Erst Start, dann End: Outlook verschiebt beim Setzen von Start das Ende mit.
	if err := a.put("Start", start); err != nil {
		return err
	}
	return a.put("End", end)
}

// setDetails: Betreff, Ort, Text, Anzeigen als, Erinnerung — nur, was
// uebergeben wurde.
func setDetails(a obj, p map[string]any) error {
	for _, f := range [][2]string{{"subject", "Subject"}, {"location", "Location"}, {"body", "Body"}} {
		if pHas(p, f[0]) {
			if err := a.put(f[1], pStr(p, f[0])); err != nil {
				return err
			}
		}
	}
	if pHas(p, "show_as") {
		code, ok := showAsCodes[strings.ToLower(pStr(p, "show_as"))]
		if !ok {
			return fmt.Errorf("show_as must be free, tentative, busy, oof or elsewhere")
		}
		if err := a.put("BusyStatus", code); err != nil {
			return err
		}
	}
	if pHas(p, "reminder") {
		m := pInt(p, "reminder", 15)
		if err := a.put("ReminderSet", m >= 0); err != nil {
			return err
		}
		if m >= 0 {
			if err := a.put("ReminderMinutesBeforeStart", m); err != nil {
				return err
			}
		}
	}
	return nil
}

// calendarCreate legt einen Termin an. Mit Teilnehmern wird es eine
// Besprechung, die gespeichert, aber NICHT verschickt wird.
func (s *session) calendarCreate(p map[string]any) error {
	a, err := callObj(s.ol, "CreateItem", olAppointmentItem)
	if err != nil {
		return err
	}
	if err := setTimes(a, p, true); err != nil {
		return err
	}
	if err := setDetails(a, p); err != nil {
		return err
	}
	req, opt := pArr(p, "attendees"), pArr(p, "optional_attendees")
	var unresolved []string
	if len(req)+len(opt) > 0 {
		if err := a.put("MeetingStatus", olMeeting); err != nil {
			return err
		}
		rs, err := getObj(a, "Recipients")
		if err != nil {
			return err
		}
		for _, l := range []struct {
			list []string
			kind int
		}{{req, olRequired}, {opt, olOptional}} {
			for _, addr := range l.list {
				if addr = strings.TrimSpace(addr); addr == "" {
					continue
				}
				r, err := callObj(rs, "Add", addr)
				if err != nil {
					return err
				}
				if err := r.put("Type", l.kind); err != nil {
					return err
				}
			}
		}
		rs.call("ResolveAll")
		each(rs, func(r obj) (bool, error) {
			if field(r, "Resolved") != true {
				unresolved = append(unresolved, str(field(r, "Name")))
			}
			return false, nil
		})
	}
	if _, err := a.call("Save"); err != nil {
		return err
	}
	o := apptInfo(a, false)
	if len(req)+len(opt) > 0 {
		o["not_sent"] = true
	}
	if len(unresolved) > 0 {
		o["unresolved"] = strings.Join(unresolved, "; ")
	}
	s.emit(o)
	return nil
}

// invited: Hat der Organisator die Einladungen schon verschickt? Unlesbar
// zaehlt als verschickt — lieber einmal zu vorsichtig.
func invited(a obj) bool {
	pa, err := getObj(a, "PropertyAccessor")
	if err != nil {
		return true
	}
	v, err := pa.call("GetProperty", prFInvited)
	if err != nil {
		return true
	}
	return v == true
}

func (s *session) calendarUpdate(p map[string]any) error {
	a, err := s.resolveAppt(pStr(p, "id"))
	if err != nil {
		return err
	}
	if pHas(p, "start") || pHas(p, "end") || pHas(p, "duration") || pHas(p, "all_day") {
		if err := setTimes(a, p, false); err != nil {
			return err
		}
	}
	if err := setDetails(a, p); err != nil {
		return err
	}
	if _, err := a.call("Save"); err != nil {
		return err
	}
	o := apptInfo(a, false)
	switch num(field(a, "MeetingStatus")) {
	case olMeeting:
		if invited(a) {
			o["note"] = "attendees are not notified — the user must open the meeting in Outlook and send the update"
		} else {
			o["not_sent"] = true
		}
	case olMeetingReceived:
		o["note"] = "only your own copy changed — the organizer's meeting is unchanged"
	}
	s.emit(o)
	return nil
}

// calendarDelete loescht Termine. Eine verschickte Besprechung, die man selbst
// organisiert, bleibt stehen: Loeschen ohne Absage liesse die Teilnehmer mit
// einem Termin zurueck, den es nicht mehr gibt, und absagen hiesse senden.
func (s *session) calendarDelete(p map[string]any) error {
	return s.forEachID(p, func(id string) error {
		a, err := s.resolveAppt(id)
		if err != nil {
			return err
		}
		subj := str(field(a, "Subject"))
		status := num(field(a, "MeetingStatus"))
		if status == olMeeting && invited(a) {
			return fmt.Errorf("you organize this meeting and invitations were sent — cancel it in Outlook so the attendees are notified")
		}
		if _, err := a.call("Delete"); err != nil {
			return err
		}
		o := map[string]any{"id": id, "subject": subj, "deleted": true}
		if status == olMeetingReceived {
			o["note"] = "the organizer is not notified — use calendar_respond with decline to tell them"
		}
		s.emit(o)
		return nil
	})
}

// calendarRespond beantwortet eine erhaltene Einladung. Respond(…, true)
// liefert die Antwort als MeetingItem, verschickt sie aber nicht; sie wird
// als Entwurf gespeichert, den der Benutzer sendet.
func (s *session) calendarRespond(p map[string]any) error {
	code, ok := responseCodes[strings.ToLower(pStr(p, "response"))]
	if !ok {
		return fmt.Errorf("response must be accept, tentative or decline")
	}
	a, err := s.resolveAppt(pStr(p, "id"))
	if err != nil {
		return err
	}
	if num(field(a, "MeetingStatus")) != olMeetingReceived {
		return fmt.Errorf("not a meeting invitation you received — only invitations from others can be answered")
	}
	r, err := callObj(a, "Respond", code, true, false)
	if err != nil {
		return err
	}
	if c := pStr(p, "comment"); c != "" {
		if err := r.put("Body", c); err != nil {
			return err
		}
	}
	if _, err := r.call("Save"); err != nil {
		return err
	}
	s.emit(map[string]any{
		"id": str(field(r, "EntryID")), "subject": str(field(r, "Subject")),
		"response": nameOf(responseNames, code), "not_sent": true,
	})
	return nil
}

// ---- Frei/Gebucht ----

// calendarFreeBusy: Recipient.FreeBusy liefert ab Mitternacht des Datums
// eine Ziffer pro Zeitfenster (0 frei, 1 mit Vorbehalt, 2 gebucht,
// 3 abwesend, 4 woanders taetig). Ausgegeben werden die belegten Bloecke
// des einen Tages.
func (s *session) calendarFreeBusy(p map[string]any) error {
	day := today()
	if d := pStr(p, "date"); d != "" {
		t, _, err := parseTime(d)
		if err != nil {
			return err
		}
		day = time.Date(t.Year(), t.Month(), t.Day(), 0, 0, 0, 0, time.UTC)
	}
	const slot = 30
	for _, who := range pArr(p, "emails") {
		if who = strings.TrimSpace(who); who == "" {
			continue
		}
		err := func() error {
			r, err := callObj(s.ns, "CreateRecipient", who)
			if err != nil {
				return err
			}
			r.call("Resolve")
			v, err := r.call("FreeBusy", day, slot, true)
			if err != nil {
				return err
			}
			fb := str(v)
			n := min(len(fb), 24*60/slot)
			var blocks []map[string]any
			for i := 0; i < n; {
				c := fb[i]
				j := i
				for j < n && fb[j] == c {
					j++
				}
				if c != '0' {
					blocks = append(blocks, map[string]any{
						"start":  isoOut(day.Add(time.Duration(i*slot) * time.Minute)),
						"end":    isoOut(day.Add(time.Duration(j*slot) * time.Minute)),
						"status": nameOf(showAsNames, int(c-'0')),
					})
				}
				i = j
			}
			s.emit(map[string]any{"email": who, "name": str(field(r, "Name")), "date": day.Format("2006-01-02"), "busy": blocks})
			return nil
		}()
		if err != nil {
			if isTransient(err) {
				return err
			}
			s.emit(map[string]any{"email": who, "error": err.Error()})
		}
	}
	return nil
}
