package main

import (
	"strings"
	"testing"

	"github.com/mark3labs/mcp-go/mcp"
)

// newReq baut einen Tool-Request mit den gegebenen Argumenten, so wie der
// MCP-Server ihn nach dem JSON-Decoding an die Handler gibt.
func newReq(args map[string]any) mcp.CallToolRequest {
	return mcp.CallToolRequest{Params: mcp.CallToolParams{Arguments: args}}
}

func TestCollectIDs(t *testing.T) {
	tests := []struct {
		name string
		args map[string]any
		want []string
		err  bool
	}{
		{"nur id (alter Aufruf)", map[string]any{"id": "AAA"}, []string{"AAA"}, false},
		{"nur ids", map[string]any{"ids": []any{"AAA", "BBB"}}, []string{"AAA", "BBB"}, false},
		{"beides, dedupliziert", map[string]any{"id": "AAA", "ids": []any{"AAA", "BBB"}}, []string{"AAA", "BBB"}, false},
		{"leere Strings raus", map[string]any{"ids": []any{"AAA", "", "  "}}, []string{"AAA"}, false},
		{"nichts angegeben", map[string]any{}, nil, true},
		{"leeres ids-Array", map[string]any{"ids": []any{}}, nil, true},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got, err := collectIDs(newReq(tc.args))
			if tc.err {
				if err == nil {
					t.Fatalf("Fehler erwartet, bekam %v", got)
				}
				return
			}
			if err != nil {
				t.Fatalf("unerwarteter Fehler: %v", err)
			}
			if strings.Join(got, ",") != strings.Join(tc.want, ",") {
				t.Errorf("got %v, want %v", got, tc.want)
			}
		})
	}
}

func TestClampTop(t *testing.T) {
	for _, tc := range []struct{ in, want int }{{0, 25}, {-5, 25}, {10, 10}, {200, 200}, {5000, 200}} {
		if got := clampTop(tc.in); got != tc.want {
			t.Errorf("clampTop(%d) = %d, want %d", tc.in, got, tc.want)
		}
	}
}

func TestParseJSONLinesSkipsGarbage(t *testing.T) {
	in := []byte(`{"id":"AAA","subject":"Erste"}` + "\n" +
		`kein json` + "\n\n" +
		`{"id":"BBB","subject":"Zweite"}` + "\n")
	got := parseJSONLines[mailItem](in)
	if len(got) != 2 || got[0].Subject != "Erste" || got[1].Subject != "Zweite" {
		t.Fatalf("kaputte Zeile nicht uebersprungen: %+v", got)
	}
}
