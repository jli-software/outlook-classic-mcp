package main

import (
	"context"
	"fmt"
	"os"
	"strings"

	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
)

// version wird bei jeder Aenderung hochgezaehlt; der Release-Workflow prueft,
// dass der Tag dazu passt.
const version = "1.0.0"

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
		mcp.WithArray("attachments", mcp.WithStringItems(), mcp.Description("Optional: local file paths to attach (e.g. a screenshot).")),
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
		mcp.WithString("dir", mcp.Description("Optional: target folder. Empty = %TEMP%\\outlook-classic-mcp-attachments.")),
	), handleSaveAttachment)

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
		fmt.Fprintf(&b, "Subject: %s\nFrom: %s\nTo: %s\nDate: %s",
			m.Subject, m.From, m.To, m.Received)
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
