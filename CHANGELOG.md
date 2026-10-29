# Changelog

All notable changes to this project are documented here. Versions follow
[Semantic Versioning](https://semver.org/).

## [1.2.0] - 2026-10-29

### Added

- Calendar tools for the default calendar:
  - `calendar_list`: appointments overlapping a time range (default: today to +7 days),
    recurring occurrences included, optionally filtered by words in subject/location.
  - `calendar_get` (batch): description, attendees with SMTP address and response,
    Teams join link, reminder, categories.
  - `calendar_create`: appointments and all-day events; with attendees a meeting that is
    saved but not sent.
  - `calendar_update`: change only the fields you pass; moving the start keeps the
    duration; works on a single occurrence of a series.
  - `calendar_delete` (batch): appointments or single occurrences. Refuses meetings you
    organize whose invitations were sent, so attendees are never left without a
    cancellation.
  - `calendar_respond`: accept, tentative or decline an invitation — saved as a draft,
    not sent.
  - `calendar_free_busy`: busy blocks of colleagues on a given day.
- Times are ISO 8601 in local time, with the local offset in the output.

## [1.1.0] - 2026-10-28

### Added

- `contacts_search`: find people in your own contacts (name, company, email) and in
  the company directory (Global Address List, by name, any word order). The directory
  is read once and kept for 30 minutes, so later searches are instant.
- `contacts_resolve`: turn a name into an SMTP address. Returns the candidates instead
  of guessing when a name is ambiguous.
- `contacts_get` (batch) and `contacts_list`: contact and directory details — phones,
  department, job title, office, address, manager.
- `outlook_whoami`: signed-in mailbox, Outlook version, connection state, and whether
  sending and which attachment folders are allowed.
- `mail_read` shows the sender's SMTP address, also for internal Exchange senders.

### Security

- Path sandbox: `mail_reply` attachments and `mail_save_attachment` folders must be
  under the user profile or `%TEMP%`. Opt out with `OUTLOOK_MCP_ALLOW_ANY_PATH=1`.
- Attachment file names from incoming mail can no longer point outside the target folder.

## [1.0.0] - 2026-10-28

First public release.

- MCP server for classic Outlook for Windows via COM — no Graph API, no app
  registration, no tokens.
- 15 tools: list, search (one folder or all folders), read, create / update / reply
  drafts, attachments (list, save), folders (list, create), move, mark read,
  categorize, delete (to Deleted Items) and sync.
- No send tool: everything outgoing is a draft.
- Reply drafts keep Outlook's HTML formatting, quote and signature;
  `mail_update_draft` replaces only your own text.
- Batch operations: several message IDs in one call; one bad ID doesn't fail the rest.
- `install` registers the server in Claude Desktop, Claude Code and Codex.
- Single EXE in pure Go: one COM thread, automatic retry when Outlook is busy,
  releases Outlook after 5 minutes idle.
