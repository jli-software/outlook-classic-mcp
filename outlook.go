package main

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"
)

// ---- Datentypen ----

type mailItem struct {
	ID          string `json:"id"`
	Subject     string `json:"subject"`
	From        string `json:"from"`
	FromEmail   string `json:"from_email"`
	To          string `json:"to"`
	Received    string `json:"received"`
	Preview     string `json:"preview"`
	Body        string `json:"body"`
	Attachments string `json:"attachments"`
	Folder      string `json:"folder"`
	Unread      bool   `json:"unread"`
	Error       string `json:"error"`
}

type folderInfo struct {
	ID     string `json:"id"`
	Path   string `json:"path"`
	Name   string `json:"name"`
	Items  int    `json:"items"`
	Unread int    `json:"unread"`
}

type actionResult struct {
	ID         string `json:"id"`
	Subject    string `json:"subject"`
	Path       string `json:"path"`
	Categories string `json:"categories"`
	Unread     bool   `json:"unread"`
	Deleted    bool   `json:"deleted"`
	Error      string `json:"error"`
}

type attachmentInfo struct {
	ID       string `json:"id"`
	Subject  string `json:"subject"`
	Index    int    `json:"index"`
	Filename string `json:"filename"`
	Size     int    `json:"size"`
	Type     int    `json:"type"`
	Empty    bool   `json:"empty"`
	Error    string `json:"error"`
}

type savedAttachment struct {
	ID       string `json:"id"`
	Subject  string `json:"subject"`
	Filename string `json:"filename"`
	Path     string `json:"path"`
	Size     int    `json:"size"`
	Empty    bool   `json:"empty"`
	Error    string `json:"error"`
}

// contactInfo ist ein eigener Kontakt oder ein Eintrag im Firmen-Adressbuch
// (Source "contacts" bzw. "directory"). Eine Zeile nur mit Note ist ein
// Hinweis zur Suche selbst.
type contactInfo struct {
	ID         string `json:"id"`
	Name       string `json:"name"`
	Email      string `json:"email"`
	Email2     string `json:"email2"`
	Email3     string `json:"email3"`
	Company    string `json:"company"`
	Department string `json:"department"`
	JobTitle   string `json:"job_title"`
	Office     string `json:"office"`
	Phone      string `json:"phone"`
	Mobile     string `json:"mobile"`
	HomePhone  string `json:"home_phone"`
	Address    string `json:"address"`
	Manager    string `json:"manager"`
	Notes      string `json:"notes"`
	Kind       string `json:"kind"`
	Source     string `json:"source"`
	Resolved   bool   `json:"resolved"`
	Note       string `json:"note"`
	Error      string `json:"error"`
}

type whoamiInfo struct {
	Name           string `json:"name"`
	Email          string `json:"email"`
	Store          string `json:"store"`
	Accounts       int    `json:"accounts"`
	Connection     string `json:"connection"`
	OutlookVersion string `json:"outlook_version"`
}

// calendarItem ist ein Termin, ein Vorkommen einer Serie (ID mit "|Start"),
// das Ergebnis einer Schreib-Op oder eine Fehlerzeile.
type calendarItem struct {
	ID         string     `json:"id"`
	Subject    string     `json:"subject"`
	Start      string     `json:"start"`
	End        string     `json:"end"`
	AllDay     bool       `json:"all_day"`
	Location   string     `json:"location"`
	ShowAs     string     `json:"show_as"`
	Meeting    string     `json:"meeting"`
	Organizer  string     `json:"organizer"`
	Response   string     `json:"response"`
	Recurring  bool       `json:"recurring"`
	Body       string     `json:"body"`
	Categories string     `json:"categories"`
	Reminder   *int       `json:"reminder"`
	TeamsURL   string     `json:"teams_url"`
	Attendees  []attendee `json:"attendees"`
	NotSent    bool       `json:"not_sent"`
	Unresolved string     `json:"unresolved"`
	Deleted    bool       `json:"deleted"`
	Note       string     `json:"note"`
	Error      string     `json:"error"`
}

type attendee struct {
	Name     string `json:"name"`
	Email    string `json:"email"`
	Type     string `json:"type"`
	Response string `json:"response"`
}

type freeBusy struct {
	Email string `json:"email"`
	Name  string `json:"name"`
	Date  string `json:"date"`
	Busy  []struct {
		Start  string `json:"start"`
		End    string `json:"end"`
		Status string `json:"status"`
	} `json:"busy"`
	Error string `json:"error"`
}

type syncResult struct {
	Synced      bool `json:"synced"`
	InboxItems  int  `json:"inbox_items"`
	InboxUnread int  `json:"inbox_unread"`
}

// parseJSONLines liest die Ergebniszeilen einer Op. Eine kaputte
// Zeile wird uebersprungen und reisst nicht den ganzen Request mit.
func parseJSONLines[T any](data []byte) []T {
	var out []T
	for _, line := range strings.Split(string(data), "\n") {
		line = strings.TrimSpace(line)
		if line == "" {
			continue
		}
		var v T
		if json.Unmarshal([]byte(line), &v) == nil {
			out = append(out, v)
		}
	}
	return out
}

// runOutlook fuehrt eine Op auf dem STA-Thread aus (sta.go, ops.go) und
// liefert ihre Ausgabe als JSON Lines.
func runOutlook(ctx context.Context, op string, params map[string]any) ([]byte, error) {
	return outlook.call(ctx, op, params)
}

// ---- Lesen & Suchen ----

func listMessages(ctx context.Context, folder string, top int, withBody bool) ([]mailItem, error) {
	out, err := runOutlook(ctx, "list", map[string]any{
		"folder": folder, "top": top, "include_body": withBody,
	})
	if err != nil {
		return nil, err
	}
	return parseJSONLines[mailItem](out), nil
}

func searchMessages(ctx context.Context, folder, query string, top int, withBody bool) ([]mailItem, error) {
	out, err := runOutlook(ctx, "search", map[string]any{
		"folder": folder, "query": query, "top": top, "include_body": withBody,
	})
	if err != nil {
		return nil, err
	}
	return parseJSONLines[mailItem](out), nil
}

func readMessages(ctx context.Context, ids []string) ([]mailItem, error) {
	out, err := runOutlook(ctx, "read", map[string]any{"ids": ids})
	if err != nil {
		return nil, err
	}
	items := parseJSONLines[mailItem](out)
	if len(items) == 0 {
		return nil, fmt.Errorf("no message found with this ID")
	}
	return items, nil
}

// ---- Entwuerfe ----

func createDraft(ctx context.Context, subject, body string, to, cc, bcc []string) (*mailItem, error) {
	params := map[string]any{"subject": subject, "body": body, "to": to}
	if len(cc) > 0 {
		params["cc"] = cc
	}
	if len(bcc) > 0 {
		params["bcc"] = bcc
	}
	return oneItem(ctx, "draft", params, "could not create draft")
}

// updateDraft erwartet in params bereits nur die Felder, die der Aufrufer
// wirklich uebergeben hat — in der Op entscheidet die Anwesenheit des
// Schluessels darueber, ob das Feld angefasst wird.
func updateDraft(ctx context.Context, params map[string]any) (*mailItem, error) {
	return oneItem(ctx, "update_draft", params, "could not update draft")
}

func replyMessage(ctx context.Context, id, body string, replyAll, keepAttach bool, attachments []string) (*mailItem, error) {
	params := map[string]any{"id": id, "body": body}
	if replyAll {
		params["reply_all"] = true
	}
	if keepAttach {
		params["keep_attachments"] = true
	}
	if len(attachments) > 0 {
		params["attachments"] = attachments
	}
	return oneItem(ctx, "reply", params, "could not create reply draft")
}

func forwardMessage(ctx context.Context, id, body string, to, cc, bcc []string) (*mailItem, error) {
	params := map[string]any{"id": id, "body": body, "to": to}
	if len(cc) > 0 {
		params["cc"] = cc
	}
	if len(bcc) > 0 {
		params["bcc"] = bcc
	}
	return oneItem(ctx, "forward", params, "could not create forward draft")
}

// sentItem: was mail_send_draft verschickt hat, Empfaenger als SMTP.
type sentItem struct {
	ID      string   `json:"id"`
	Subject string   `json:"subject"`
	To      []string `json:"to"`
	CC      []string `json:"cc"`
	BCC     []string `json:"bcc"`
	Sent    bool     `json:"sent"`
}

func sendDraft(ctx context.Context, id string) (*sentItem, error) {
	out, err := runOutlook(ctx, "send_draft", map[string]any{"id": id})
	if err != nil {
		return nil, err
	}
	rs := parseJSONLines[sentItem](out)
	if len(rs) == 0 || !rs[0].Sent {
		return nil, fmt.Errorf("no confirmation from Outlook — check the Outbox before trying again")
	}
	return &rs[0], nil
}

func oneItem(ctx context.Context, op string, params map[string]any, failMsg string) (*mailItem, error) {
	out, err := runOutlook(ctx, op, params)
	if err != nil {
		return nil, err
	}
	items := parseJSONLines[mailItem](out)
	if len(items) == 0 {
		return nil, fmt.Errorf("%s", failMsg)
	}
	return &items[0], nil
}

// ---- Sync ----

// syncMailbox stoesst ein Senden/Empfangen an und wartet wait Sekunden, damit
// der OST-Cache frisch wird (wichtig, wenn das klassische Outlook selten laeuft).
func syncMailbox(ctx context.Context, wait int) (*syncResult, error) {
	params := map[string]any{}
	if wait > 0 {
		params["wait"] = wait
	}
	out, err := runOutlook(ctx, "sync", params)
	if err != nil {
		return nil, err
	}
	rs := parseJSONLines[syncResult](out)
	if len(rs) == 0 {
		return nil, fmt.Errorf("no answer from sync")
	}
	return &rs[0], nil
}

// ---- Ordner & Verwaltung ----

func listFolders(ctx context.Context) ([]folderInfo, error) {
	out, err := runOutlook(ctx, "folders", map[string]any{})
	if err != nil {
		return nil, err
	}
	return parseJSONLines[folderInfo](out), nil
}

func createFolder(ctx context.Context, parentID, name string) (*folderInfo, error) {
	out, err := runOutlook(ctx, "create_folder", map[string]any{
		"parent_id": parentID, "name": name,
	})
	if err != nil {
		return nil, err
	}
	fs := parseJSONLines[folderInfo](out)
	if len(fs) == 0 {
		return nil, fmt.Errorf("could not create folder")
	}
	return &fs[0], nil
}

func actions(ctx context.Context, op string, params map[string]any) ([]actionResult, error) {
	out, err := runOutlook(ctx, op, params)
	if err != nil {
		return nil, err
	}
	rs := parseJSONLines[actionResult](out)
	if len(rs) == 0 {
		return nil, fmt.Errorf("no answer from Outlook")
	}
	return rs, nil
}

func moveMessages(ctx context.Context, ids []string, destID string) ([]actionResult, error) {
	return actions(ctx, "move", map[string]any{"ids": ids, "folder_id": destID})
}

func markRead(ctx context.Context, ids []string, read bool) ([]actionResult, error) {
	return actions(ctx, "mark", map[string]any{"ids": ids, "read": read})
}

func categorize(ctx context.Context, ids []string, categories []string) ([]actionResult, error) {
	return actions(ctx, "categorize", map[string]any{"ids": ids, "categories": categories})
}

func deleteMessages(ctx context.Context, ids []string) ([]actionResult, error) {
	return actions(ctx, "delete", map[string]any{"ids": ids})
}

// ---- Anhaenge ----

func listAttachments(ctx context.Context, ids []string) ([]attachmentInfo, error) {
	out, err := runOutlook(ctx, "attachments", map[string]any{"ids": ids})
	if err != nil {
		return nil, err
	}
	return parseJSONLines[attachmentInfo](out), nil
}

func saveAttachments(ctx context.Context, ids []string, dir string, index int) ([]savedAttachment, error) {
	params := map[string]any{"ids": ids}
	if dir != "" {
		params["dir"] = dir
	}
	if index > 0 {
		params["index"] = index
	}
	out, err := runOutlook(ctx, "save_attachment", params)
	if err != nil {
		return nil, err
	}
	return parseJSONLines[savedAttachment](out), nil
}

// ---- Kontakte ----

func contacts(ctx context.Context, op string, params map[string]any) ([]contactInfo, error) {
	out, err := runOutlook(ctx, op, params)
	if err != nil {
		return nil, err
	}
	return parseJSONLines[contactInfo](out), nil
}

func searchContacts(ctx context.Context, query string, top int) ([]contactInfo, error) {
	return contacts(ctx, "contacts_search", map[string]any{"query": query, "top": top})
}

func resolveContact(ctx context.Context, name string) ([]contactInfo, error) {
	return contacts(ctx, "contacts_resolve", map[string]any{"name": name})
}

func getContacts(ctx context.Context, ids []string) ([]contactInfo, error) {
	return contacts(ctx, "contacts_get", map[string]any{"ids": ids})
}

func listContacts(ctx context.Context, top int) ([]contactInfo, error) {
	return contacts(ctx, "contacts_list", map[string]any{"top": top})
}

func whoami(ctx context.Context) (*whoamiInfo, error) {
	out, err := runOutlook(ctx, "whoami", map[string]any{})
	if err != nil {
		return nil, err
	}
	rs := parseJSONLines[whoamiInfo](out)
	if len(rs) == 0 {
		return nil, fmt.Errorf("no answer from Outlook")
	}
	return &rs[0], nil
}

// ---- Kalender ----

func calendar(ctx context.Context, op string, params map[string]any) ([]calendarItem, error) {
	out, err := runOutlook(ctx, op, params)
	if err != nil {
		return nil, err
	}
	return parseJSONLines[calendarItem](out), nil
}

func freeBusyOf(ctx context.Context, emails []string, date string) ([]freeBusy, error) {
	out, err := runOutlook(ctx, "calendar_free_busy", map[string]any{"emails": emails, "date": date})
	if err != nil {
		return nil, err
	}
	return parseJSONLines[freeBusy](out), nil
}
