<div align="center">
  <img src="logo.svg" width="120" alt="Outlook Classic MCP logo" />
  <h1>Outlook Classic MCP</h1>
  <p><b>An MCP server for Microsoft Outlook — the classic desktop app for Windows.</b><br/>
  Let Claude, Codex and other AI assistants read, search and organize your mailbox,<br/>
  look up colleagues and write drafts for you — through the Outlook that is already signed in on your PC.</p>

  [![Release](https://img.shields.io/github/v/release/jli-software/outlook-classic-mcp?style=flat-square&color=4cc2ff)](https://github.com/jli-software/outlook-classic-mcp/releases/latest)
  ![Platform](https://img.shields.io/badge/platform-Windows-89b4fa?style=flat-square)
  ![Outlook](https://img.shields.io/badge/Outlook-classic%20(Win32)-89b4fa?style=flat-square)
  ![License](https://img.shields.io/badge/license-MIT-a6e3a1?style=flat-square)
  ![Built with Go](https://img.shields.io/badge/built%20with-Go-f5a97f?style=flat-square)
</div>

---

```
Claude / Codex  ──MCP──▶  outlook-classic-mcp.exe  ──COM──▶  Outlook (classic)  ──▶  your mailbox
```

## Why

- **No Azure app, no Graph API, no tokens.** It talks to your running Outlook via COM and
  uses the session you are already signed in with. Nothing to register, no admin consent.
- **Works where Microsoft Graph is blocked** — e.g. by Conditional Access or token
  protection policies. If Outlook works on your PC, this works.
- **Local only.** Your mail never leaves your machine except to the AI client you use.
- **Never sends.** Everything outgoing becomes a draft you review and send yourself.
- **Files stay fenced in.** Attachments can only be added from — and saved to — your user
  profile and `%TEMP%`, so a crafted email can't talk the assistant into attaching
  arbitrary files.
- **One small EXE.** No PowerShell, no runtime, no background service.

## Requirements

- Windows with **classic Outlook** (Outlook for Microsoft 365 / 2016 / 2019 / 2021 desktop,
  the Win32 app) set up with your mailbox.
  The **new Outlook** (“Outlook for Windows”, olk.exe) has no COM interface —
  see [Using the new Outlook?](#using-the-new-outlook) below.
- An MCP client: Claude Desktop, Claude Code, Codex, Cursor, VS Code, …
- Run it as your normal user, **not elevated** (an admin process cannot attach to your Outlook).

## Install

In PowerShell:

```powershell
irm https://raw.githubusercontent.com/jli-software/outlook-classic-mcp/main/install.ps1 | iex
```

This downloads the latest release to `%LOCALAPPDATA%\Programs\outlook-classic-mcp`,
checks its SHA-256 and registers it as server `outlook` in **Claude Desktop**,
**Claude Code** and **Codex**. Restart your clients — done.

<details>
<summary>Manual install</summary>

1. Download `outlook-classic-mcp-windows-x64.exe` from the
   [latest release](https://github.com/jli-software/outlook-classic-mcp/releases/latest)
   and save it as `outlook-classic-mcp.exe` wherever you like.
2. Register it:
   ```powershell
   .\outlook-classic-mcp.exe install           # Claude Desktop, Claude Code, Codex
   .\outlook-classic-mcp.exe install --print   # dry run: only show the config files
   .\outlook-classic-mcp.exe install codex     # only one target (desktop | code | codex)
   ```
   Existing configs are kept; a `.outlook-classic-mcp.bak` backup is written first.

Windows SmartScreen may warn about the unsigned EXE — “More info → Run anyway”.
</details>

<details>
<summary>Other MCP clients (Cursor, VS Code, Windsurf, …)</summary>

Add a stdio server to your client's MCP config:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\Users\\you\\AppData\\Local\\Programs\\outlook-classic-mcp\\outlook-classic-mcp.exe",
      "args": ["serve"]
    }
  }
}
```
</details>

## What you can ask

> “Summarize my unread mail from today.”
> “Find the invoice from Contoso in all folders and save the PDF.”
> “Draft a reply to Anna's last mail: Thursday works, 10am.”
> “Move all newsletters from this week to Archive and mark them read.”
> “What's Anna Smith's email and department? Draft her an invite for Friday.”

## Tools

| Tool | What it does |
|---|---|
| `mail_list` | Newest emails of a folder (Inbox, Sent Items, Drafts, … or any folder) |
| `mail_search` | Search subject, sender and body — one folder or `folder: "all"` |
| `mail_read` | Full emails with body and attachment names (several at once) |
| `mail_create_draft` | New draft with To / CC / BCC |
| `mail_reply` | Reply draft with thread and quote, optionally with attachments |
| `mail_update_draft` | Change a draft; in replies only your text is replaced, quote and signature stay |
| `mail_attachments` | List attachments |
| `mail_save_attachment` | Save attachments to disk so the assistant can read them |
| `mail_folders` | Complete folder tree with IDs and counts |
| `mail_create_folder` | Create a folder |
| `mail_move` | Move emails to a folder |
| `mail_mark_read` | Mark read / unread |
| `mail_categorize` | Set Outlook categories (color tags) |
| `mail_delete` | Move to Deleted Items (never deletes permanently) |
| `mail_sync` | Trigger Send/Receive to refresh the local cache |
| `contacts_search` | Find people in your contacts and the company directory (Global Address List) |
| `contacts_resolve` | Name → email address; lists the candidates instead of guessing when ambiguous |
| `contacts_get` | All details of a contact or directory entry (phones, department, manager, …) |
| `contacts_list` | Your own contacts |
| `outlook_whoami` | Signed-in mailbox, Outlook version, connection state and this server's settings |

There is deliberately **no send tool**.

Email addresses always come back as SMTP addresses (`anna.smith@contoso.com`), also for
Exchange senders and directory entries — never as `/O=…/CN=…`.

### Attachment folders

`mail_reply` (`attachments`) and `mail_save_attachment` (`dir`) only accept paths under
your user profile (`%USERPROFILE%`) or `%TEMP%`; anything else is refused. To lift this,
set `OUTLOOK_MCP_ALLOW_ANY_PATH=1` in the server's environment, e.g.

```json
"outlook": { "command": "…\\outlook-classic-mcp.exe", "args": ["serve"],
             "env": { "OUTLOOK_MCP_ALLOW_ANY_PATH": "1" } }
```

## Using the new Outlook?

You can keep working in the new Outlook. The classic Outlook only has to be installed
and should run now and then, because this server reads its local cache:

- **Best:** start classic Outlook minimized with Windows. You won't notice it, and the
  cache stays fresh. Turn off its desktop notifications so you don't get every mail twice
  (*File → Options → Mail → Display a Desktop Alert*).
  ```powershell
  $s  = New-Object -ComObject WScript.Shell
  $sc = $s.CreateShortcut("$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\Outlook (classic).lnk")
  $sc.TargetPath  = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE').'(default)'
  $sc.WindowStyle = 7   # minimized
  $sc.Save()
  ```
- **On demand:** ask the assistant to run `mail_sync` first.

If classic Outlook is closed, the server starts it invisibly and releases it again
after 5 minutes without a call.

## Troubleshooting

| Problem | Fix |
|---|---|
| `0x80080005` / “server execution failed” | Your MCP client runs elevated, or Outlook is starting. Run the client as normal user; keep Outlook open. |
| Drafts don't show up in the new Outlook | Classic Outlook hasn't synced yet — run `mail_sync` or keep classic Outlook running. |
| `folder not found` | Use `mail_folders` and pass the folder's ID or full path. Unknown folders are an error on purpose, never a silent fallback to the Inbox. |
| First `contacts_search` takes a few seconds | Outlook can't filter the company directory, so the first search reads all names (~2 ms per entry) and keeps them for 30 minutes. Later searches are instant. |
| Search over all folders takes seconds | It scans the text of every message (~5–10 s for a few thousand). Search one folder when you can. |
| A dialog in Outlook blocks | Calls time out after 90 s; close the dialog and try again. |

## Build from source

```powershell
git clone https://github.com/jli-software/outlook-classic-mcp
cd outlook-classic-mcp
go build -o outlook-classic-mcp.exe .
go test .
```

Pure Go (no CGO), using [go-ole](https://github.com/go-ole/go-ole) for COM and
[mcp-go](https://github.com/mark3labs/mcp-go) for MCP.

## Disclaimer

This is an independent open-source project. It is not affiliated with, endorsed by or
sponsored by Microsoft. “Microsoft” and “Outlook” are trademarks of the Microsoft group
of companies.

## License

[MIT](LICENSE)
