package main

import (
	"fmt"
	"regexp"
	"strings"
)

// Antworten und Entwuerfe im HTML-Format. Nie .Body setzen: das stellt die
// ganze Mail auf Nur-Text um und zerschiesst Zitat, Links und Signatur
// (siehe AGENTS.md).

var (
	reWordSection = regexp.MustCompile(`(?i)<div class="?WordSection1"?[^>]*>`)
	reBodyTag     = regexp.MustCompile(`(?i)<body[^>]*>`)
	reMarker      = regexp.MustCompile(`(?i)<div id="?mailmcp-text"?[^>]*>`)

	// Ende des eigenen Texts: der Absatz mit dem Word-Lesezeichen der Signatur
	// bzw. des Schreibbereichs, oder der Zitatkopf — grauer Trennrahmen
	// (Outlook Classic) bzw. divRplyFwdMsg (OWA), jeweils samt der unmittelbar
	// davor geoeffneten Huellen-<div>s. RE2 kennt kein Lookahead: ein Tag, das
	// nicht mit "/p" beginnt, ist hier <(?:[^/>][^>]*|/[^p>][^>]*|/)>.
	reDraftEnd = []*regexp.Regexp{
		regexp.MustCompile(`(?i)<p[^>]*>(?:\s*<(?:[^/>][^>]*|/[^p>][^>]*|/)>)*?\s*<a name="?_MailAutoSig`),
		regexp.MustCompile(`(?i)<p[^>]*>(?:\s*<(?:[^/>][^>]*|/[^p>][^>]*|/)>)*?\s*<a name="?_MailEndCompose`),
		regexp.MustCompile(`(?i)(?:<div[^>]*>\s*)*<div[^>]*border-top:solid #E1E1E1`),
		regexp.MustCompile(`(?i)(?:<hr[^>]*>\s*)?(?:<div[^>]*>\s*)*<div[^>]*id="?divRplyFwdMsg`),
	}
)

// textHTML macht aus dem eigenen Text einen HTML-Block. Absaetze im Word-Stil
// (MsoNormal), damit Outlook die Standardschrift nimmt. Der Container mit
// id=mailmcp-text markiert den eigenen Text, damit update_draft spaeter genau
// ihn ersetzen kann (Outlook behaelt die id beim Speichern per COM, live
// geprueft).
func textHTML(text string) string {
	var b strings.Builder
	b.WriteString(`<div id=mailmcp-text>`)
	for _, line := range strings.Split(strings.ReplaceAll(text, "\r\n", "\n"), "\n") {
		if strings.TrimSpace(line) == "" {
			b.WriteString(`<p class=MsoNormal>&nbsp;</p>`)
		} else {
			b.WriteString(`<p class=MsoNormal>` + htmlEncode(line) + `</p>`)
		}
	}
	b.WriteString(`<p class=MsoNormal>&nbsp;</p></div>`)
	return b.String()
}

// htmlEncode kodiert wie .NET WebUtility.HtmlEncode:
// die fuenf Sonderzeichen, Latin-1 ab U+00A0 und alles ausserhalb der BMP als
// numerische Referenz.
func htmlEncode(s string) string {
	var b strings.Builder
	for _, r := range s {
		switch {
		case r == '<':
			b.WriteString("&lt;")
		case r == '>':
			b.WriteString("&gt;")
		case r == '&':
			b.WriteString("&amp;")
		case r == '"':
			b.WriteString("&quot;")
		case r == '\'':
			b.WriteString("&#39;")
		case (r >= 0xA0 && r <= 0xFF) || r >= 0x10000:
			fmt.Fprintf(&b, "&#%d;", r)
		default:
			b.WriteRune(r)
		}
	}
	return b.String()
}

// bodyStart: Position direkt hinter dem Beginn des Textbereichs — innerhalb
// von WordSection1, sonst ausserhalb der Formatierung; ohne WordSection1
// hinter <body>, ohne beides ganz vorne.
func bodyStart(html string) int {
	if loc := reWordSection.FindStringIndex(html); loc != nil {
		return loc[1]
	}
	if loc := reBodyTag.FindStringIndex(html); loc != nil {
		return loc[1]
	}
	return 0
}

// insertReplyHTML setzt den eigenen Text vor das Zitat einer Antwort.
func insertReplyHTML(html, text string) string {
	i := bodyStart(html)
	return html[:i] + textHTML(text) + html[i:]
}

// replaceDraftHTML ersetzt den eigenen Text eines HTML-Entwurfs, ohne Zitat
// und Signatur anzufassen:
//  1. Markierter Container (von insertReplyHTML) -> genau ihn ersetzen.
//  2. Sonst alles vom Textbeginn bis zur Signatur bzw. zum Zitatkopf (Entwurf
//     aus Outlook selbst, oder Markierung beim Bearbeiten in Outlook verloren).
//  3. Ohne Signatur und Zitat den ganzen Inhalt von WordSection1 bzw. <body>.
func replaceDraftHTML(html, text string) string {
	block := textHTML(text)
	low := asciiLower(html)

	// Der Container zaehlt nur, wenn vor seinem </div> kein weiteres <div
	// beginnt — ein verschachtelter Container waere nicht sauber abgrenzbar.
	for _, loc := range reMarker.FindAllStringIndex(html, -1) {
		c := strings.Index(low[loc[1]:], "</div>")
		if c < 0 || strings.Contains(low[loc[1]:loc[1]+c], "<div") {
			continue
		}
		end := loc[1] + c + len("</div>")
		return html[:loc[0]] + block + html[end:]
	}

	start := bodyStart(html)
	end := -1
	for _, re := range reDraftEnd {
		if loc := re.FindStringIndex(html[start:]); loc != nil {
			if i := start + loc[0]; end < 0 || i < end {
				end = i
			}
		}
	}
	if end < 0 {
		// Kein Zitat: bis zum schliessenden </div> von WordSection1 bzw. </body>.
		end = strings.LastIndex(low, "</body>")
		if end < 0 {
			end = len(html)
		}
		if end > 0 && strings.Contains(low[:start], "wordsection1") {
			if d := strings.LastIndex(low[:end], "</div>"); d >= start {
				end = d
			}
		}
	}
	return html[:start] + block + html[end:]
}

// asciiLower senkt nur A–Z ab und laesst damit jede Byte-Position stehen —
// strings.ToLower darf das nicht, Indizes muessen auf html passen.
func asciiLower(s string) string {
	b := []byte(s)
	for i, c := range b {
		if c >= 'A' && c <= 'Z' {
			b[i] = c + 32
		}
	}
	return string(b)
}
