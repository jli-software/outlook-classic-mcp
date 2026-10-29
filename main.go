package main

import (
	"context"
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
)

// version wird bei jeder Aenderung hochgezaehlt; der Release-Workflow prueft,
// dass der Tag dazu passt.
const version = "1.2.0"

func main() {
	if len(os.Args) > 1 {
		switch os.Args[1] {
		case "install", "config":
			runInstall(os.Args[2:])
			return
		case "serve":
			// faellt durch zu runServer()
		case "version", "--version", "-v":
			fmt.Println("outlook-classic-mcp", version)
			return
		case "-h", "--help", "help":
			printUsage()
			return
		default:
			fmt.Fprintf(os.Stderr, "Unknown command %q.\n\n", os.Args[1])
			printUsage()
			os.Exit(2)
		}
	}
	// Standard: MCP-Server ueber stdio. Keine Auth noetig — wir nutzen das lokal
	// angemeldete Outlook per COM (kein Token, kein Conditional Access).
	runServer()
}

func printUsage() {
	fmt.Println(`outlook-classic-mcp ` + version + ` — MCP server for classic Outlook for Windows (local, via COM)

Commands:
  serve              Start the MCP server (default, called by your MCP client)
  install            Register in Claude Desktop, Claude Code and Codex
  install --print    Show what would be registered (dry run)
  install codex      Only one target (desktop | code | codex)
  version            Print the version
  help               This help`)
}

func runServer() {
	// Der STA-Thread startet lazy beim ersten Tool-Call und gibt beim
	// Herunterfahren Outlook frei, damit keine COM-Referenz zurueckbleibt.
	defer outlook.stop()

	s := server.NewMCPServer("outlook-classic-mcp", version)

	folderDesc := "Folder: alias (inbox = default, sentitems, drafts, deleteditems, junk), OR an EntryID from mail_folders, OR a path ('\\Inbox\\Projects' or 'Inbox/Projects'), OR a unique folder name. An unknown folder is an error, never an empty result."
	idDesc := "EntryID of one message (from mail_list/mail_search)."
	idsDesc := "Several EntryIDs at once — ONE Outlook round trip instead of one per message. Prefer this whenever several messages are involved."
	bodyDesc := "true = include the full text of every message instead of a 160-character preview. Saves a later mail_read but costs a lot of context — only useful with a small 'top'."

	s.AddTool(mcp.NewTool("mail_list",
		mcp.WithDescription("Lists the newest emails in an Outlook folder (default: Inbox). Returns ID, subject, sender, date, preview and read status."),
		mcp.WithString("folder", mcp.Description(folderDesc)),
		mcp.WithNumber("top", mcp.Description("Number of messages (default 25, max 200).")),
		mcp.WithBoolean("include_body", mcp.Description(bodyDesc)),
	), handleList)

	s.AddTool(mcp.NewTool("mail_search",
		mcp.WithDescription("Searches one folder (default Inbox) or, with folder='all', every mail folder of the mailbox by subject, sender or body text. With 'all' every hit shows its folder; results are sorted by date across all folders (newest first) and 'top' applies to the total."),
		mcp.WithString("query", mcp.Required(), mcp.Description("Search term, e.g. 'invoice' or 'Smith'.")),
		mcp.WithString("folder", mcp.Description(folderDesc+" Additionally 'all' = every mail folder, including Deleted Items and Junk.")),
		mcp.WithNumber("top", mcp.Description("Number of hits (default 25, max 200).")),
		mcp.WithBoolean("include_body", mcp.Description(bodyDesc)),
	), handleSearch)

	s.AddTool(mcp.NewTool("mail_read",
		mcp.WithDescription("Reads one or more complete emails including body text by ID. Several IDs in ONE call are much faster than several calls."),
		mcp.WithString("id", mcp.Description(idDesc)),
		mcp.WithArray("ids", mcp.WithStringItems(), mcp.Description(idsDesc)),
	), handleRead)

	s.AddTool(mcp.NewTool("mail_create_draft",
		mcp.WithDescription("Creates an email DRAFT in Outlook. It is NOT sent — the user reviews and sends it."),
		mcp.WithString("subject", mcp.Required(), mcp.Description("Subject.")),
		mcp.WithString("body", mcp.Required(), mcp.Description("Message text (plain text).")),
		mcp.WithArray("to", mcp.Required(), mcp.WithStringItems(), mcp.Description("Recipient addresses (To).")),
		mcp.WithArray("cc", mcp.WithStringItems(), mcp.Description("Optional: CC addresses.")),
		mcp.WithArray("bcc", mcp.WithStringItems(), mcp.Description("Optional: BCC addresses.")),
	), handleCreateDraft)

	s.AddTool(mcp.NewTool("mail_update_draft",
		mcp.WithDescription("Changes an existing DRAFT. Only the fields you pass are changed (ideal for corrections instead of creating a new draft)."),
		mcp.WithString("id", mcp.Required(), mcp.Description("EntryID of the draft.")),
		mcp.WithString("subject", mcp.Description("New subject.")),
		mcp.WithString("body", mcp.Description("New message text (plain text). In reply drafts only your own text is replaced; quote, signature and HTML formatting stay.")),
		mcp.WithArray("to", mcp.WithStringItems(), mcp.Description("New To addresses (replaces).")),
		mcp.WithArray("cc", mcp.WithStringItems(), mcp.Description("New CC addresses (replaces).")),
		mcp.WithArray("bcc", mcp.WithStringItems(), mcp.Description("New BCC addresses (replaces).")),
	), handleUpdateDraft)

	s.AddTool(mcp.NewTool("mail_reply",
		mcp.WithDescription("Creates a REPLY draft to an existing email (thread and quote of the original added automatically, RE: subject, recipients prefilled). It is NOT sent."),
		mcp.WithString("id", mcp.Required(), mcp.Description("EntryID of the original email.")),
		mcp.WithString("body", mcp.Required(), mcp.Description("Your reply text (placed above the quote).")),
		mcp.WithBoolean("reply_all", mcp.Description("true = reply to all (including CC). Default false.")),
		mcp.WithBoolean("keep_attachments", mcp.Description("true = carry over the original's attachments. Default false.")),
		mcp.WithArray("attachments", mcp.WithStringItems(), mcp.Description("Optional: local file paths to attach (e.g. a screenshot). Only files under the user profile or %TEMP% (see outlook_whoami).")),
	), handleReply)

	s.AddTool(mcp.NewTool("mail_sync",
		mcp.WithDescription("Triggers Send/Receive in classic Outlook and refreshes the local cache. Useful BEFORE reading or searching when classic Outlook rarely runs (e.g. because the user works in the new Outlook). Returns the Inbox item and unread counts afterwards."),
		mcp.WithNumber("wait", mcp.Description("Seconds to wait so the sync can finish (default 8, max 60).")),
	), handleSync)

	s.AddTool(mcp.NewTool("mail_folders",
		mcp.WithDescription("Lists the mailbox's complete folder tree with path, EntryID, item count and unread count. Use the ID as target for mail_move / mail_create_folder."),
	), handleFolders)

	s.AddTool(mcp.NewTool("mail_create_folder",
		mcp.WithDescription("Creates a new folder. Without parent_id it is created under the Inbox."),
		mcp.WithString("name", mcp.Required(), mcp.Description("Name of the new folder.")),
		mcp.WithString("parent_id", mcp.Description("EntryID of the parent folder (from mail_folders). Empty = Inbox.")),
	), handleCreateFolder)

	s.AddTool(mcp.NewTool("mail_move",
		mcp.WithDescription("Moves one or more emails into a folder."),
		mcp.WithString("id", mcp.Description(idDesc)),
		mcp.WithArray("ids", mcp.WithStringItems(), mcp.Description(idsDesc)),
		mcp.WithString("folder_id", mcp.Required(), mcp.Description("EntryID of the target folder (from mail_folders).")),
	), handleMove)

	s.AddTool(mcp.NewTool("mail_mark_read",
		mcp.WithDescription("Marks one or more emails as read or unread."),
		mcp.WithString("id", mcp.Description(idDesc)),
		mcp.WithArray("ids", mcp.WithStringItems(), mcp.Description(idsDesc)),
		mcp.WithBoolean("read", mcp.Description("true = read (default), false = unread.")),
	), handleMarkRead)

	s.AddTool(mcp.NewTool("mail_categorize",
		mcp.WithDescription("Sets Outlook categories (color tags) on one or more emails. Replaces existing categories."),
		mcp.WithString("id", mcp.Description(idDesc)),
		mcp.WithArray("ids", mcp.WithStringItems(), mcp.Description(idsDesc)),
		mcp.WithArray("categories", mcp.Required(), mcp.WithStringItems(), mcp.Description("Category names, e.g. ['Important','Waiting']. An empty array removes all.")),
	), handleCategorize)

	s.AddTool(mcp.NewTool("mail_delete",
		mcp.WithDescription("Moves one or more emails to Deleted Items (never deletes permanently)."),
		mcp.WithString("id", mcp.Description(idDesc)),
		mcp.WithArray("ids", mcp.WithStringItems(), mcp.Description(idsDesc)),
	), handleDelete)

	s.AddTool(mcp.NewTool("mail_attachments",
		mcp.WithDescription("Lists the attachments of one or more emails (index, file name, size in bytes)."),
		mcp.WithString("id", mcp.Description(idDesc)),
		mcp.WithArray("ids", mcp.WithStringItems(), mcp.Description(idsDesc)),
	), handleAttachments)

	s.AddTool(mcp.NewTool("mail_save_attachment",
		mcp.WithDescription("Saves attachments of one or more emails to local disk and returns the file paths (readable afterwards with normal file reading)."),
		mcp.WithString("id", mcp.Description(idDesc)),
		mcp.WithArray("ids", mcp.WithStringItems(), mcp.Description(idsDesc)),
		mcp.WithNumber("index", mcp.Description("Optional: only this attachment (1-based index from mail_attachments). Empty = all. Applies to every given email.")),
		mcp.WithString("dir", mcp.Description("Optional: target folder under the user profile or %TEMP%. Empty = %TEMP%\\outlook-classic-mcp-attachments.")),
	), handleSaveAttachment)

	s.AddTool(mcp.NewTool("contacts_search",
		mcp.WithDescription("Searches your Outlook contacts (name, company, email) and the company directory (Global Address List, by name) for people and distribution lists. Returns name, SMTP email, company, department, phone and an ID for contacts_get. Word order doesn't matter ('Smith Anna' finds 'Anna Smith')."),
		mcp.WithString("query", mcp.Required(), mcp.Description("Name or part of it, company or email, e.g. 'Smith' or 'Anna Sm'.")),
		mcp.WithNumber("top", mcp.Description("Maximum hits per source (contacts, directory). Default 10, max 50.")),
	), handleContactsSearch)

	s.AddTool(mcp.NewTool("contacts_resolve",
		mcp.WithDescription("Turns a person's name into their email (SMTP) address — use this before writing a draft to someone you only know by name. Returns exactly one address, or the list of candidates if the name is ambiguous (never guesses)."),
		mcp.WithString("name", mcp.Required(), mcp.Description("Name, e.g. 'Anna Smith', or an email address to check.")),
	), handleContactsResolve)

	s.AddTool(mcp.NewTool("contacts_get",
		mcp.WithDescription("Shows all details of one or more contacts or directory entries by ID (from contacts_search/contacts_list): emails, phones, company, department, job title, office, address, manager, notes."),
		mcp.WithString("id", mcp.Description("ID of one contact or directory entry.")),
		mcp.WithArray("ids", mcp.WithStringItems(), mcp.Description("Several IDs at once — one Outlook round trip.")),
	), handleContactsGet)

	s.AddTool(mcp.NewTool("contacts_list",
		mcp.WithDescription("Lists your own Outlook contacts (not the company directory), sorted by name."),
		mcp.WithNumber("top", mcp.Description("Number of contacts (default 50, max 200).")),
	), handleContactsList)

	s.AddTool(mcp.NewTool("outlook_whoami",
		mcp.WithDescription("Shows the signed-in mailbox (name, email), Outlook version and connection state, and this server's settings: whether sending is possible and which folders attachments may be read from or saved to."),
	), handleWhoami)

	timeDesc := "ISO 8601 in local time, e.g. '2026-10-29T14:00' (an offset like '+01:00' or 'Z' is converted to local time)."
	calIDDesc := "Calendar item ID from calendar_list. An ID with '|' is one occurrence of a series; without it, the whole series or a single appointment."
	s.AddTool(mcp.NewTool("calendar_list",
		mcp.WithDescription("Lists appointments and meetings in the default calendar that overlap a time range (default: today to +7 days), recurring occurrences included, sorted by start. Returns ID, time, subject, location, show-as, organizer and your response."),
		mcp.WithString("start", mcp.Description("Range start. "+timeDesc+" A date alone = midnight. Default: today 00:00.")),
		mcp.WithString("end", mcp.Description("Range end. A date alone counts to the end of that day. Default: start + 7 days.")),
		mcp.WithString("query", mcp.Description("Optional: only items whose subject or location contains all these words.")),
		mcp.WithNumber("top", mcp.Description("Maximum number of items (default 50, max 200).")),
	), handleCalendarList)

	s.AddTool(mcp.NewTool("calendar_get",
		mcp.WithDescription("Shows full details of one or more calendar items: description, attendees with their responses, Teams join link, reminder, categories."),
		mcp.WithString("id", mcp.Description(calIDDesc)),
		mcp.WithArray("ids", mcp.WithStringItems(), mcp.Description("Several IDs at once — one Outlook round trip.")),
	), handleCalendarGet)

	s.AddTool(mcp.NewTool("calendar_create",
		mcp.WithDescription("Creates an appointment in the default calendar. With attendees it becomes a meeting that is SAVED BUT NOT SENT — the user opens it in Outlook and clicks Send. Resolve names to addresses with contacts_resolve first."),
		mcp.WithString("subject", mcp.Required(), mcp.Description("Subject.")),
		mcp.WithString("start", mcp.Required(), mcp.Description("Start. "+timeDesc+" A date alone ('2026-10-29') makes an all-day event.")),
		mcp.WithString("end", mcp.Description("End (same format). For all-day events a date alone is the last day, inclusive. Default: start + duration.")),
		mcp.WithNumber("duration", mcp.Description("Length in minutes when no end is given (default 30; all-day: one day).")),
		mcp.WithBoolean("all_day", mcp.Description("true = all-day event.")),
		mcp.WithString("location", mcp.Description("Location.")),
		mcp.WithString("body", mcp.Description("Description (plain text).")),
		mcp.WithArray("attendees", mcp.WithStringItems(), mcp.Description("Required attendees (email addresses). Makes it a meeting — not sent.")),
		mcp.WithArray("optional_attendees", mcp.WithStringItems(), mcp.Description("Optional attendees (email addresses).")),
		mcp.WithString("show_as", mcp.Description("free, tentative, busy (default), oof or elsewhere.")),
		mcp.WithNumber("reminder", mcp.Description("Reminder in minutes before start; -1 = no reminder. Default: Outlook's setting.")),
	), handleCalendarCreate)

	s.AddTool(mcp.NewTool("calendar_update",
		mcp.WithDescription("Changes a calendar item. Only the fields you pass are changed; moving the start keeps the duration. For a series pass an occurrence ID to change only that occurrence. Attendees are NOT notified — for a meeting you organize the user sends the update in Outlook."),
		mcp.WithString("id", mcp.Required(), mcp.Description(calIDDesc)),
		mcp.WithString("subject", mcp.Description("New subject.")),
		mcp.WithString("start", mcp.Description("New start. "+timeDesc)),
		mcp.WithString("end", mcp.Description("New end.")),
		mcp.WithNumber("duration", mcp.Description("New length in minutes (instead of end).")),
		mcp.WithBoolean("all_day", mcp.Description("true/false = make it an all-day event or not.")),
		mcp.WithString("location", mcp.Description("New location.")),
		mcp.WithString("body", mcp.Description("New description (plain text). Replaces the whole description, including any Teams join information — leave it out to keep it.")),
		mcp.WithString("show_as", mcp.Description("free, tentative, busy, oof or elsewhere.")),
		mcp.WithNumber("reminder", mcp.Description("Reminder in minutes before start; -1 = no reminder.")),
	), handleCalendarUpdate)

	s.AddTool(mcp.NewTool("calendar_delete",
		mcp.WithDescription("Deletes one or more calendar items (to Deleted Items). An occurrence ID deletes only that occurrence. A meeting you organize whose invitations were sent is refused (cancelling means sending — the user cancels it in Outlook). Deleting a received meeting does not notify the organizer — use calendar_respond with decline for that."),
		mcp.WithString("id", mcp.Description(calIDDesc)),
		mcp.WithArray("ids", mcp.WithStringItems(), mcp.Description("Several IDs at once.")),
	), handleCalendarDelete)

	s.AddTool(mcp.NewTool("calendar_respond",
		mcp.WithDescription("Answers a meeting invitation you received: accept, tentative or decline. The response is saved as a DRAFT and NOT sent — the user sends it from Drafts."),
		mcp.WithString("id", mcp.Required(), mcp.Description(calIDDesc)),
		mcp.WithString("response", mcp.Required(), mcp.Description("accept, tentative or decline.")),
		mcp.WithString("comment", mcp.Description("Optional message to the organizer.")),
	), handleCalendarRespond)

	s.AddTool(mcp.NewTool("calendar_free_busy",
		mcp.WithDescription("Shows when people are busy on one day (free/busy from Exchange, in 30-minute blocks) — e.g. to find a time for a meeting. Works for colleagues in your organization; external addresses usually have no data."),
		mcp.WithArray("emails", mcp.Required(), mcp.WithStringItems(), mcp.Description("Email addresses (resolve names with contacts_resolve first).")),
		mcp.WithString("date", mcp.Description("Day, e.g. '2026-10-29'. Default: today.")),
	), handleCalendarFreeBusy)

	if err := server.ServeStdio(s); err != nil {
		fmt.Fprintln(os.Stderr, "Server error:", err)
		os.Exit(1)
	}
}

func clampTop(t int) int {
	if t <= 0 {
		return 25
	}
	if t > 200 {
		return 200
	}
	return t
}

// collectIDs nimmt 'id' (einzeln) und 'ids' (Batch) — beides gleichzeitig ist
// erlaubt, Duplikate fliegen raus.
func collectIDs(r mcp.CallToolRequest) ([]string, error) {
	seen := map[string]bool{}
	var ids []string
	add := func(s string) {
		if s = strings.TrimSpace(s); s != "" && !seen[s] {
			seen[s] = true
			ids = append(ids, s)
		}
	}
	add(r.GetString("id", ""))
	for _, s := range r.GetStringSlice("ids", nil) {
		add(s)
	}
	if len(ids) == 0 {
		return nil, fmt.Errorf("either 'id' (one message) or 'ids' (several) is required")
	}
	return ids, nil
}

func formatList(msgs []mailItem) string {
	if len(msgs) == 0 {
		return "No messages found."
	}
	var b strings.Builder
	for i, m := range msgs {
		mark := " "
		if m.Unread {
			mark = "•" // ungelesen
		}
		fmt.Fprintf(&b, "%d. %s [%s] %s\n   From: %s | %s\n   ID: %s\n",
			i+1, mark, m.Received, m.Subject, m.From, strings.TrimSpace(m.Preview), m.ID)
		if m.Folder != "" {
			fmt.Fprintf(&b, "   Folder: %s\n", m.Folder)
		}
		if strings.TrimSpace(m.Body) != "" {
			fmt.Fprintf(&b, "   ---\n%s\n   ---\n", m.Body)
		}
	}
	return b.String()
}

// formatActions fasst das Ergebnis einer Batch-Operation zusammen. Ein
// fehlgeschlagener Eintrag darf den Rest nicht verdecken.
func formatActions(rs []actionResult, line func(actionResult) string) string {
	var b strings.Builder
	ok, fail := 0, 0
	for _, a := range rs {
		if a.Error != "" {
			fail++
			fmt.Fprintf(&b, "✗ %s — %s\n", a.ID, a.Error)
			continue
		}
		ok++
		b.WriteString(line(a) + "\n")
	}
	if len(rs) > 1 {
		fmt.Fprintf(&b, "\n%d succeeded, %d failed.\n", ok, fail)
	}
	return b.String()
}

func handleList(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	msgs, err := listMessages(ctx, r.GetString("folder", ""), clampTop(r.GetInt("top", 25)),
		r.GetBool("include_body", false))
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_list failed", err), nil
	}
	return mcp.NewToolResultText(formatList(msgs)), nil
}

func handleSearch(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	query, err := r.RequireString("query")
	if err != nil {
		return mcp.NewToolResultErrorFromErr("parameter 'query' is missing", err), nil
	}
	msgs, err := searchMessages(ctx, r.GetString("folder", ""), query,
		clampTop(r.GetInt("top", 25)), r.GetBool("include_body", false))
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_search failed", err), nil
	}
	return mcp.NewToolResultText(formatList(msgs)), nil
}

func handleRead(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	ids, err := collectIDs(r)
	if err != nil {
		return mcp.NewToolResultError(err.Error()), nil
	}
	items, err := readMessages(ctx, ids)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_read failed", err), nil
	}
	var b strings.Builder
	for i, m := range items {
		if i > 0 {
			b.WriteString("\n" + strings.Repeat("─", 60) + "\n\n")
		}
		if m.Error != "" {
			fmt.Fprintf(&b, "✗ %s — %s\n", m.ID, m.Error)
			continue
		}
		from := m.From
		if m.FromEmail != "" && !strings.EqualFold(m.FromEmail, m.From) {
			from += " <" + m.FromEmail + ">"
		}
		fmt.Fprintf(&b, "Subject: %s\nFrom: %s\nTo: %s\nDate: %s",
			m.Subject, from, m.To, m.Received)
		if strings.TrimSpace(m.Attachments) != "" {
			b.WriteString("\nAttachments: " + m.Attachments)
		}
		b.WriteString("\nID: " + m.ID + "\n\n" + m.Body + "\n")
	}
	return mcp.NewToolResultText(b.String()), nil
}

func handleCreateDraft(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	subject, err := r.RequireString("subject")
	if err != nil {
		return mcp.NewToolResultErrorFromErr("parameter 'subject' is missing", err), nil
	}
	body, err := r.RequireString("body")
	if err != nil {
		return mcp.NewToolResultErrorFromErr("parameter 'body' is missing", err), nil
	}
	to, err := r.RequireStringSlice("to")
	if err != nil || len(to) == 0 {
		return mcp.NewToolResultError("parameter 'to' is missing or empty."), nil
	}
	m, err := createDraft(ctx, subject, body, to,
		r.GetStringSlice("cc", nil), r.GetStringSlice("bcc", nil))
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_create_draft failed", err), nil
	}
	return mcp.NewToolResultText(fmt.Sprintf("Draft created in Outlook (not sent).\nID: %s\nSubject: %s", m.ID, m.Subject)), nil
}

func handleUpdateDraft(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	id, err := r.RequireString("id")
	if err != nil {
		return mcp.NewToolResultErrorFromErr("parameter 'id' is missing", err), nil
	}
	// Nur die tatsaechlich uebergebenen Felder weitergeben — in der Op
	// entscheidet die Anwesenheit des Schluessels, ob das Feld angefasst wird.
	args := r.GetArguments()
	params := map[string]any{"id": id}
	for _, f := range []string{"subject", "body"} {
		if _, ok := args[f]; ok {
			params[f] = r.GetString(f, "")
		}
	}
	for _, f := range []string{"to", "cc", "bcc"} {
		if _, ok := args[f]; ok {
			v := r.GetStringSlice(f, nil)
			if v == nil {
				v = []string{}
			}
			params[f] = v
		}
	}
	if len(params) == 1 {
		return mcp.NewToolResultError("Nothing to update — pass at least one field (subject/body/to/cc/bcc)."), nil
	}
	m, err := updateDraft(ctx, params)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_update_draft failed", err), nil
	}
	return mcp.NewToolResultText(fmt.Sprintf("Draft updated.\nID: %s\nSubject: %s", m.ID, m.Subject)), nil
}

func handleReply(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	id, err := r.RequireString("id")
	if err != nil {
		return mcp.NewToolResultErrorFromErr("parameter 'id' is missing", err), nil
	}
	body, err := r.RequireString("body")
	if err != nil {
		return mcp.NewToolResultErrorFromErr("parameter 'body' is missing", err), nil
	}
	m, err := replyMessage(ctx, id, body, r.GetBool("reply_all", false),
		r.GetBool("keep_attachments", false), r.GetStringSlice("attachments", nil))
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_reply failed", err), nil
	}
	return mcp.NewToolResultText(fmt.Sprintf("Reply draft created (not sent).\nID: %s\nSubject: %s", m.ID, m.Subject)), nil
}

func handleFolders(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	folders, err := listFolders(ctx)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_folders failed", err), nil
	}
	if len(folders) == 0 {
		return mcp.NewToolResultText("No folders found."), nil
	}
	var b strings.Builder
	for _, f := range folders {
		fmt.Fprintf(&b, "%s  (%d items, %d unread)\n   ID: %s\n", f.Path, f.Items, f.Unread, f.ID)
	}
	return mcp.NewToolResultText(b.String()), nil
}

func handleSync(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	wait := r.GetInt("wait", 8)
	if wait < 0 {
		wait = 0
	}
	if wait > 60 {
		wait = 60
	}
	s, err := syncMailbox(ctx, wait)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_sync failed", err), nil
	}
	return mcp.NewToolResultText(fmt.Sprintf("Sync triggered. Inbox: %d items, %d unread.", s.InboxItems, s.InboxUnread)), nil
}

func handleCreateFolder(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	name, err := r.RequireString("name")
	if err != nil {
		return mcp.NewToolResultErrorFromErr("parameter 'name' is missing", err), nil
	}
	f, err := createFolder(ctx, r.GetString("parent_id", ""), name)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_create_folder failed", err), nil
	}
	return mcp.NewToolResultText(fmt.Sprintf("Folder created: %s\nID: %s", f.Path, f.ID)), nil
}

func handleMove(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	ids, err := collectIDs(r)
	if err != nil {
		return mcp.NewToolResultError(err.Error()), nil
	}
	folderID, err := r.RequireString("folder_id")
	if err != nil {
		return mcp.NewToolResultErrorFromErr("parameter 'folder_id' is missing", err), nil
	}
	rs, err := moveMessages(ctx, ids, folderID)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_move failed", err), nil
	}
	return mcp.NewToolResultText(formatActions(rs, func(a actionResult) string {
		return fmt.Sprintf("✓ Moved to %s: %s\n   New ID: %s", a.Path, a.Subject, a.ID)
	})), nil
}

func handleMarkRead(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	ids, err := collectIDs(r)
	if err != nil {
		return mcp.NewToolResultError(err.Error()), nil
	}
	rs, err := markRead(ctx, ids, r.GetBool("read", true))
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_mark_read failed", err), nil
	}
	return mcp.NewToolResultText(formatActions(rs, func(a actionResult) string {
		status := "read"
		if a.Unread {
			status = "unread"
		}
		return fmt.Sprintf("✓ Marked as %s: %s", status, a.Subject)
	})), nil
}

func handleCategorize(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	ids, err := collectIDs(r)
	if err != nil {
		return mcp.NewToolResultError(err.Error()), nil
	}
	cats := r.GetStringSlice("categories", nil)
	if cats == nil {
		cats = []string{}
	}
	rs, err := categorize(ctx, ids, cats)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_categorize failed", err), nil
	}
	return mcp.NewToolResultText(formatActions(rs, func(a actionResult) string {
		return fmt.Sprintf("✓ Categories [%s]: %s", a.Categories, a.Subject)
	})), nil
}

func handleDelete(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	ids, err := collectIDs(r)
	if err != nil {
		return mcp.NewToolResultError(err.Error()), nil
	}
	rs, err := deleteMessages(ctx, ids)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_delete failed", err), nil
	}
	return mcp.NewToolResultText(formatActions(rs, func(a actionResult) string {
		return fmt.Sprintf("✓ Moved to Deleted Items: %s", a.Subject)
	})), nil
}

// mailHeader beschriftet die Anhang-Gruppe einer Mail. Bei mehreren IDs muss
// erkennbar bleiben, welcher Anhang zu welcher Nachricht gehoert.
func mailHeader(b *strings.Builder, cur *string, id, subject string) {
	if *cur == id {
		return
	}
	if *cur != "" {
		b.WriteString("\n")
	}
	if subject == "" {
		subject = "(unknown subject)"
	}
	fmt.Fprintf(b, "%s\n   ID: %s\n", subject, id)
	*cur = id
}

func handleAttachments(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	ids, err := collectIDs(r)
	if err != nil {
		return mcp.NewToolResultError(err.Error()), nil
	}
	atts, err := listAttachments(ctx, ids)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_attachments failed", err), nil
	}
	if len(atts) == 0 {
		return mcp.NewToolResultText("No attachments."), nil
	}
	var b strings.Builder
	cur := ""
	for _, a := range atts {
		mailHeader(&b, &cur, a.ID, a.Subject)
		switch {
		case a.Error != "":
			fmt.Fprintf(&b, "   ✗ %s\n", a.Error)
		case a.Empty:
			b.WriteString("   (no attachments)\n")
		default:
			fmt.Fprintf(&b, "   %d. %s  (%d bytes)\n", a.Index, a.Filename, a.Size)
		}
	}
	return mcp.NewToolResultText(b.String()), nil
}

func handleSaveAttachment(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	ids, err := collectIDs(r)
	if err != nil {
		return mcp.NewToolResultError(err.Error()), nil
	}
	saved, err := saveAttachments(ctx, ids, r.GetString("dir", ""), r.GetInt("index", 0))
	if err != nil {
		return mcp.NewToolResultErrorFromErr("mail_save_attachment failed", err), nil
	}
	if len(saved) == 0 {
		return mcp.NewToolResultText("No attachments found to save."), nil
	}
	var b strings.Builder
	cur := ""
	for _, s := range saved {
		mailHeader(&b, &cur, s.ID, s.Subject)
		switch {
		case s.Error != "":
			fmt.Fprintf(&b, "   ✗ %s — %s\n", s.Filename, s.Error)
		case s.Empty:
			b.WriteString("   (no attachments)\n")
		default:
			fmt.Fprintf(&b, "   ✓ %s  (%d bytes)\n      %s\n", s.Filename, s.Size, s.Path)
		}
	}
	return mcp.NewToolResultText(b.String()), nil
}

// ---- Kontakte ----

// formatContacts listet Personen; Zeilen nur mit Note kommen als Hinweis ans Ende.
func formatContacts(cs []contactInfo, full bool) string {
	var b strings.Builder
	n := 0
	for _, c := range cs {
		if c.Note != "" && c.ID == "" && c.Name == "" {
			continue
		}
		if c.Error != "" {
			fmt.Fprintf(&b, "✗ %s — %s\n", c.ID, c.Error)
			continue
		}
		n++
		name := c.Name
		if c.Email != "" {
			name += " <" + c.Email + ">"
		}
		if c.Kind == "list" {
			name += "  (distribution list)"
		}
		fmt.Fprintf(&b, "%d. %s\n", n, name)
		var org []string
		for _, v := range []string{c.JobTitle, c.Department, c.Company, c.Office} {
			if v != "" {
				org = append(org, v)
			}
		}
		if len(org) > 0 {
			fmt.Fprintf(&b, "   %s\n", strings.Join(org, " | "))
		}
		lines := [][2]string{{"Phone", c.Phone}, {"Mobile", c.Mobile}}
		if full {
			lines = append(lines, [][2]string{{"Email 2", c.Email2}, {"Email 3", c.Email3},
				{"Home phone", c.HomePhone}, {"Address", c.Address}, {"Manager", c.Manager}}...)
		}
		for _, l := range lines {
			if l[1] != "" {
				fmt.Fprintf(&b, "   %s: %s\n", l[0], l[1])
			}
		}
		src := "your contacts"
		if c.Source == "directory" {
			src = "company directory"
		}
		fmt.Fprintf(&b, "   Source: %s\n", src)
		if full && c.Notes != "" {
			fmt.Fprintf(&b, "   Notes:\n%s\n", c.Notes)
		}
		if c.ID != "" {
			fmt.Fprintf(&b, "   ID: %s\n", c.ID)
		}
	}
	if n == 0 && b.Len() == 0 {
		return ""
	}
	return b.String() + notes(cs)
}

func notes(cs []contactInfo) string {
	var b strings.Builder
	for _, c := range cs {
		if c.Note != "" {
			b.WriteString("\nNote: " + c.Note + "\n")
		}
	}
	return b.String()
}

func clampN(v, d, max int) int {
	if v <= 0 {
		return d
	}
	return min(v, max)
}

func handleContactsSearch(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	query, err := r.RequireString("query")
	if err != nil || strings.TrimSpace(query) == "" {
		return mcp.NewToolResultError("parameter 'query' is missing or empty."), nil
	}
	cs, err := searchContacts(ctx, query, clampN(r.GetInt("top", 10), 10, 50))
	if err != nil {
		return mcp.NewToolResultErrorFromErr("contacts_search failed", err), nil
	}
	if out := formatContacts(cs, false); out != "" {
		return mcp.NewToolResultText(out), nil
	}
	return mcp.NewToolResultText("No contacts found." + notes(cs)), nil
}

func handleContactsResolve(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	name, err := r.RequireString("name")
	if err != nil || strings.TrimSpace(name) == "" {
		return mcp.NewToolResultError("parameter 'name' is missing or empty."), nil
	}
	cs, err := resolveContact(ctx, name)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("contacts_resolve failed", err), nil
	}
	people := 0
	for _, c := range cs {
		if c.Resolved {
			return mcp.NewToolResultText(fmt.Sprintf("Resolved: %s <%s>\n\n%s", c.Name, c.Email,
				formatContacts([]contactInfo{c}, false))), nil
		}
		if c.Name != "" || c.ID != "" {
			people++
		}
	}
	if people == 0 {
		return mcp.NewToolResultText(fmt.Sprintf("No match for %q.\n%s", name, notes(cs))), nil
	}
	return mcp.NewToolResultText(fmt.Sprintf("%q is ambiguous — %d candidates. Ask the user which one is meant:\n\n%s",
		name, people, formatContacts(cs, false))), nil
}

func handleContactsGet(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	ids, err := collectIDs(r)
	if err != nil {
		return mcp.NewToolResultError(err.Error()), nil
	}
	cs, err := getContacts(ctx, ids)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("contacts_get failed", err), nil
	}
	if out := formatContacts(cs, true); out != "" {
		return mcp.NewToolResultText(out), nil
	}
	return mcp.NewToolResultText("No contact found."), nil
}

func handleContactsList(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	cs, err := listContacts(ctx, clampN(r.GetInt("top", 50), 50, 200))
	if err != nil {
		return mcp.NewToolResultErrorFromErr("contacts_list failed", err), nil
	}
	if out := formatContacts(cs, false); out != "" {
		return mcp.NewToolResultText(out), nil
	}
	return mcp.NewToolResultText("No contacts."), nil
}

func handleWhoami(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	w, err := whoami(ctx)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("outlook_whoami failed", err), nil
	}
	var b strings.Builder
	fmt.Fprintf(&b, "Mailbox: %s <%s>\n", w.Name, w.Email)
	if w.Store != "" && !strings.EqualFold(w.Store, w.Email) {
		fmt.Fprintf(&b, "Store: %s\n", w.Store)
	}
	if w.Accounts > 1 {
		fmt.Fprintf(&b, "Accounts in profile: %d (this server uses the default store)\n", w.Accounts)
	}
	fmt.Fprintf(&b, "Outlook: %s", w.OutlookVersion)
	if w.Connection != "" {
		fmt.Fprintf(&b, ", %s", w.Connection)
	}
	fmt.Fprintf(&b, "\nServer: outlook-classic-mcp %s\n", version)
	b.WriteString("Sending: off — everything outgoing is a draft the user sends\n")
	if sandboxOn() {
		fmt.Fprintf(&b, "Attachment paths: only under %s\n", strings.Join(allowedRoots(), ", "))
	} else {
		fmt.Fprintf(&b, "Attachment paths: any (%s=1)\n", allowAnyPathEnv)
	}
	return mcp.NewToolResultText(b.String()), nil
}

// ---- Kalender ----

// passArgs reicht nur die uebergebenen Argumente weiter — die Ops
// entscheiden ueber die Anwesenheit des Schluessels (pHas).
func passArgs(r mcp.CallToolRequest, keys ...string) map[string]any {
	args := r.GetArguments()
	p := map[string]any{}
	for _, k := range keys {
		if v, ok := args[k]; ok && v != nil {
			p[k] = v
		}
	}
	return p
}

func showTime(iso string, withTime bool) string {
	t, err := time.Parse(time.RFC3339, iso)
	if err != nil {
		return iso
	}
	if withTime {
		return t.Format("Mon 2006-01-02 15:04")
	}
	return t.Format("Mon 2006-01-02")
}

func clock(iso string) string {
	t, err := time.Parse(time.RFC3339, iso)
	if err != nil {
		return iso
	}
	return t.Format("15:04")
}

func calWhen(c calendarItem) string {
	from := showTime(c.Start, false)
	if c.AllDay {
		// Das Ende ist Mitternacht nach dem letzten Tag.
		if t, err := time.Parse(time.RFC3339, c.End); err == nil {
			if last := t.AddDate(0, 0, -1).Format("Mon 2006-01-02"); last != from {
				return from + " – " + last + ", all day"
			}
		}
		return from + ", all day"
	}
	if showTime(c.End, false) == from {
		return showTime(c.Start, true) + "–" + clock(c.End)
	}
	return showTime(c.Start, true) + " – " + showTime(c.End, true)
}

func formatCalendar(cs []calendarItem, full bool) string {
	var b strings.Builder
	n := 0
	for _, c := range cs {
		if c.Error != "" {
			fmt.Fprintf(&b, "✗ %s — %s\n", c.ID, c.Error)
			continue
		}
		n++
		if full && n > 1 {
			b.WriteString("\n" + strings.Repeat("─", 60) + "\n\n")
		}
		fmt.Fprintf(&b, "%d. %s  %s\n", n, calWhen(c), c.Subject)
		var info []string
		if c.Location != "" {
			info = append(info, "Location: "+c.Location)
		}
		info = append(info, c.ShowAs)
		switch c.Meeting {
		case "organizer":
			info = append(info, "meeting, you organize")
		case "attendee":
			info = append(info, fmt.Sprintf("meeting by %s, your response: %s", c.Organizer, c.Response))
		case "canceled":
			info = append(info, "CANCELED by "+c.Organizer)
		}
		if c.Recurring {
			info = append(info, "recurring")
		}
		fmt.Fprintf(&b, "   %s\n", strings.Join(info, " | "))
		if full {
			for _, a := range c.Attendees {
				who := a.Name
				if a.Email != "" && !strings.EqualFold(a.Email, a.Name) {
					who += " <" + a.Email + ">"
				}
				fmt.Fprintf(&b, "   Attendee (%s): %s — %s\n", a.Type, who, a.Response)
			}
			if c.TeamsURL != "" {
				fmt.Fprintf(&b, "   Teams: %s\n", c.TeamsURL)
			}
			if c.Reminder != nil && *c.Reminder >= 0 {
				fmt.Fprintf(&b, "   Reminder: %d min before\n", *c.Reminder)
			}
			if c.Categories != "" {
				fmt.Fprintf(&b, "   Categories: %s\n", c.Categories)
			}
		}
		fmt.Fprintf(&b, "   ID: %s\n", c.ID)
		if full && strings.TrimSpace(c.Body) != "" {
			fmt.Fprintf(&b, "\n%s\n", strings.TrimSpace(c.Body))
		}
	}
	return b.String()
}

// calNotes: was der Benutzer nach einer Schreib-Op wissen muss.
func calNotes(c calendarItem) string {
	var b strings.Builder
	if c.NotSent {
		b.WriteString("\nInvitations NOT sent — the user opens the meeting in Outlook and clicks Send.")
	}
	if c.Unresolved != "" {
		b.WriteString("\nNot resolved (check these addresses): " + c.Unresolved)
	}
	if c.Note != "" {
		b.WriteString("\nNote: " + c.Note)
	}
	return b.String()
}

func handleCalendarList(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	p := passArgs(r, "start", "end", "query")
	p["top"] = clampN(r.GetInt("top", 50), 50, 200)
	cs, err := calendar(ctx, "calendar_list", p)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("calendar_list failed", err), nil
	}
	if out := formatCalendar(cs, false); out != "" {
		return mcp.NewToolResultText(out), nil
	}
	return mcp.NewToolResultText("No calendar items in this range."), nil
}

func handleCalendarGet(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	ids, err := collectIDs(r)
	if err != nil {
		return mcp.NewToolResultError(err.Error()), nil
	}
	cs, err := calendar(ctx, "calendar_get", map[string]any{"ids": ids})
	if err != nil {
		return mcp.NewToolResultErrorFromErr("calendar_get failed", err), nil
	}
	if out := formatCalendar(cs, true); out != "" {
		return mcp.NewToolResultText(out), nil
	}
	return mcp.NewToolResultText("No calendar item found."), nil
}

var calFields = []string{"subject", "start", "end", "duration", "all_day", "location", "body", "show_as", "reminder"}

func handleCalendarCreate(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	if s, err := r.RequireString("subject"); err != nil || strings.TrimSpace(s) == "" {
		return mcp.NewToolResultError("parameter 'subject' is missing or empty."), nil
	}
	if _, err := r.RequireString("start"); err != nil {
		return mcp.NewToolResultError("parameter 'start' is missing."), nil
	}
	cs, err := calendar(ctx, "calendar_create", passArgs(r, append(calFields, "attendees", "optional_attendees")...))
	if err != nil {
		return mcp.NewToolResultErrorFromErr("calendar_create failed", err), nil
	}
	if len(cs) == 0 {
		return mcp.NewToolResultError("calendar_create: no answer from Outlook"), nil
	}
	head := "Appointment created."
	if cs[0].NotSent {
		head = "Meeting saved in the calendar."
	}
	return mcp.NewToolResultText(head + "\n" + formatCalendar(cs, false) + calNotes(cs[0])), nil
}

func handleCalendarUpdate(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	id, err := r.RequireString("id")
	if err != nil {
		return mcp.NewToolResultErrorFromErr("parameter 'id' is missing", err), nil
	}
	p := passArgs(r, calFields...)
	if len(p) == 0 {
		return mcp.NewToolResultError("Nothing to update — pass at least one field (" + strings.Join(calFields, "/") + ")."), nil
	}
	p["id"] = id
	cs, err := calendar(ctx, "calendar_update", p)
	if err != nil {
		return mcp.NewToolResultErrorFromErr("calendar_update failed", err), nil
	}
	if len(cs) == 0 {
		return mcp.NewToolResultError("calendar_update: no answer from Outlook"), nil
	}
	return mcp.NewToolResultText("Calendar item updated.\n" + formatCalendar(cs, false) + calNotes(cs[0])), nil
}

func handleCalendarDelete(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	ids, err := collectIDs(r)
	if err != nil {
		return mcp.NewToolResultError(err.Error()), nil
	}
	cs, err := calendar(ctx, "calendar_delete", map[string]any{"ids": ids})
	if err != nil {
		return mcp.NewToolResultErrorFromErr("calendar_delete failed", err), nil
	}
	var b strings.Builder
	ok, fail := 0, 0
	for _, c := range cs {
		if c.Error != "" {
			fail++
			fmt.Fprintf(&b, "✗ %s — %s\n", c.ID, c.Error)
			continue
		}
		ok++
		fmt.Fprintf(&b, "✓ Deleted: %s\n", c.Subject)
		if c.Note != "" {
			fmt.Fprintf(&b, "   Note: %s\n", c.Note)
		}
	}
	if len(cs) > 1 {
		fmt.Fprintf(&b, "\n%d succeeded, %d failed.\n", ok, fail)
	}
	return mcp.NewToolResultText(b.String()), nil
}

func handleCalendarRespond(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	id, err := r.RequireString("id")
	if err != nil {
		return mcp.NewToolResultErrorFromErr("parameter 'id' is missing", err), nil
	}
	resp, err := r.RequireString("response")
	if err != nil {
		return mcp.NewToolResultErrorFromErr("parameter 'response' is missing", err), nil
	}
	cs, err := calendar(ctx, "calendar_respond", map[string]any{"id": id, "response": resp, "comment": r.GetString("comment", "")})
	if err != nil {
		return mcp.NewToolResultErrorFromErr("calendar_respond failed", err), nil
	}
	if len(cs) == 0 {
		return mcp.NewToolResultError("calendar_respond: no answer from Outlook"), nil
	}
	return mcp.NewToolResultText(fmt.Sprintf("Response (%s) saved as a draft — NOT sent. The user sends it from Drafts.\nSubject: %s\nID: %s",
		cs[0].Response, cs[0].Subject, cs[0].ID)), nil
}

func handleCalendarFreeBusy(ctx context.Context, r mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	emails := r.GetStringSlice("emails", nil)
	if len(emails) == 0 {
		return mcp.NewToolResultError("parameter 'emails' is missing or empty."), nil
	}
	fbs, err := freeBusyOf(ctx, emails, r.GetString("date", ""))
	if err != nil {
		return mcp.NewToolResultErrorFromErr("calendar_free_busy failed", err), nil
	}
	var b strings.Builder
	for _, f := range fbs {
		if f.Error != "" {
			fmt.Fprintf(&b, "✗ %s — %s\n", f.Email, f.Error)
			continue
		}
		who := f.Email
		if f.Name != "" && !strings.EqualFold(f.Name, f.Email) {
			who = f.Name + " <" + f.Email + ">"
		}
		fmt.Fprintf(&b, "%s, %s:\n", who, f.Date)
		if len(f.Busy) == 0 {
			b.WriteString("   free all day\n")
		}
		for _, x := range f.Busy {
			fmt.Fprintf(&b, "   %s–%s %s\n", clock(x.Start), clock(x.End), x.Status)
		}
	}
	return mcp.NewToolResultText(b.String()), nil
}
