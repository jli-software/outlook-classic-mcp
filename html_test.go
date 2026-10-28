package main

import "testing"

// Antworttext muss als HTML vor dem Zitat landen, nicht die Mail auf Nur-Text
// umstellen — sonst zerfallen Zitat, Links und Signatur.
func TestInsertReplyHTML(t *testing.T) {
	quote := `<html><body lang=DE-CH><div class=WordSection1><p>Zitat</p></div></body></html>`
	got := insertReplyHTML(quote, "Hi Anna,\r\n\r\nPreis < 5 & \"gut\" – Grüsse")
	want := `<html><body lang=DE-CH><div class=WordSection1><div id=mailmcp-text>` +
		`<p class=MsoNormal>Hi Anna,</p><p class=MsoNormal>&nbsp;</p>` +
		`<p class=MsoNormal>Preis &lt; 5 &amp; &quot;gut&quot; – Gr&#252;sse</p>` +
		`<p class=MsoNormal>&nbsp;</p></div><p>Zitat</p></div></body></html>`
	if got != want {
		t.Errorf("HTML falsch:\n got %s\nwant %s", got, want)
	}
}

func TestBodyStartFallbacks(t *testing.T) {
	if got := insertReplyHTML(`<html><BODY class=x><p>Z</p></body></html>`, "A"); got !=
		`<html><BODY class=x>`+textHTML("A")+`<p>Z</p></body></html>` {
		t.Errorf("ohne WordSection1 nicht hinter <body>: %s", got)
	}
	if got := insertReplyHTML(`<p>Z</p>`, "A"); got != textHTML("A")+`<p>Z</p>` {
		t.Errorf("ohne <body> nicht ganz vorne: %s", got)
	}
}

func TestHTMLEncodeLikeDotNet(t *testing.T) {
	// Wie .NET WebUtility.HtmlEncode: Latin-1 numerisch,
	// der Gedankenstrich (U+2013) bleibt, Emoji (ausserhalb BMP) numerisch.
	got := htmlEncode(`'ä' – é ✓ 😀 <&>`)
	want := `&#39;&#228;&#39; – &#233; ✓ &#128512; &lt;&amp;&gt;`
	if got != want {
		t.Errorf("\n got %s\nwant %s", got, want)
	}
}

// update_draft darf bei HTML-Entwuerfen nur den eigenen Text ersetzen:
// Zitat, Signatur und Format bleiben (Falle 9).
func TestReplaceDraftHTML(t *testing.T) {
	const head = `<html><body lang=DE><div class=WordSection1>`
	const quote = `<p class=MsoNormal><a name="_MailEndCompose"><o:p>&nbsp;</o:p></a></p>` +
		`<div><div style='border:none;border-top:solid #E1E1E1 1.0pt'><p>Von: X</p></div><p>Zitat</p></div>`
	const tail = `</div></body></html>`
	text := func(s string) string {
		return `<div id=mailmcp-text><p class=MsoNormal>` + s + `</p><p class=MsoNormal>&nbsp;</p></div>`
	}

	t.Run("Marker ersetzen, zweimal", func(t *testing.T) {
		html := replaceDraftHTML(head+text("Alt")+quote+tail, "Neu")
		if want := head + text("Neu") + quote + tail; html != want {
			t.Fatalf("erstes Ersetzen:\n got %s\nwant %s", html, want)
		}
		html = replaceDraftHTML(html, "Neuer")
		if want := head + text("Neuer") + quote + tail; html != want {
			t.Fatalf("zweites Ersetzen:\n got %s\nwant %s", html, want)
		}
	})

	t.Run("Marker in Anfuehrungszeichen, wie Word speichert", func(t *testing.T) {
		in := head + `<div id="mailmcp-text"><p class=MsoNormal>Alt</p></div>` + quote + tail
		if got, want := replaceDraftHTML(in, "Neu"), head+text("Neu")+quote+tail; got != want {
			t.Fatalf("\n got %s\nwant %s", got, want)
		}
	})

	t.Run("ohne Marker bis zum Lesezeichen", func(t *testing.T) {
		in := head + `<p class=MsoNormal>Von Hand</p><p class=MsoNormal>&nbsp;</p>` + quote + tail
		if got, want := replaceDraftHTML(in, "Neu"), head+text("Neu")+quote+tail; got != want {
			t.Fatalf("\n got %s\nwant %s", got, want)
		}
	})

	t.Run("ohne Marker, nur Zitatrahmen", func(t *testing.T) {
		q := `<div><div style='border:none;border-top:solid #E1E1E1 1.0pt'><p>Von: X</p></div></div>`
		in := head + `<p>Von Hand</p>` + q + tail
		if got, want := replaceDraftHTML(in, "Neu"), head+text("Neu")+q+tail; got != want {
			t.Fatalf("\n got %s\nwant %s", got, want)
		}
	})

	t.Run("OWA-Zitatkopf", func(t *testing.T) {
		q := `<hr style="display:inline-block"><div id="divRplyFwdMsg"><p>From: X</p></div>`
		in := head + `<p>Von Hand</p>` + q + tail
		if got, want := replaceDraftHTML(in, "Neu"), head+text("Neu")+q+tail; got != want {
			t.Fatalf("\n got %s\nwant %s", got, want)
		}
	})

	t.Run("Signatur bleibt", func(t *testing.T) {
		sig := `<p class=MsoNormal><a name="_MailAutoSig"><span>Gruss Anna</span></a></p>`
		in := head + `<p>Von Hand</p>` + sig + quote + tail
		if got, want := replaceDraftHTML(in, "Neu"), head+text("Neu")+sig+quote+tail; got != want {
			t.Fatalf("\n got %s\nwant %s", got, want)
		}
	})

	t.Run("Lesezeichen erst nach </p> zaehlt nicht als Signaturabsatz", func(t *testing.T) {
		// Das <a name> steht in einem spaeteren Absatz; der erste Absatz endet
		// vorher und darf nicht mitgenommen werden.
		sig := `<p class=MsoNormal><span><a name="_MailAutoSig">Gruss</a></span></p>`
		in := head + `<p class=MsoNormal><b>Alt</b></p>` + sig + tail
		if got, want := replaceDraftHTML(in, "Neu"), head+text("Neu")+sig+tail; got != want {
			t.Fatalf("\n got %s\nwant %s", got, want)
		}
	})

	t.Run("ohne Zitat ganzer Text", func(t *testing.T) {
		in := head + `<p class=MsoNormal>Alt 1</p><p class=MsoNormal>Alt 2</p>` + tail
		if got, want := replaceDraftHTML(in, "Neu"), head+text("Neu")+tail; got != want {
			t.Fatalf("\n got %s\nwant %s", got, want)
		}
	})
}
