// Command po-distribution-extractor converts a purchase-order PDF into a
// per-branch distribution spreadsheet.
package main

import (
	"flag"
	"fmt"
	"os"
	"sort"
	"strings"

	"github.com/ledongthuc/pdf"
)

func main() {
	inFlag := flag.String("in", "", "input PDF")
	flag.Parse()

	if *inFlag == "" {
		fmt.Println("usage: po-distribution-extractor -in <file.pdf>")
		return
	}
	lines, err := extractLines(*inFlag, nil)
	if err != nil {
		fmt.Fprintln(os.Stderr, "error:", err)
		os.Exit(1)
	}
	fmt.Printf("read %d text lines from %s\n", len(lines), *inFlag)
}

// extractLines returns one reconstructed text line per visual row of the PDF.
// onPage (may be nil) is called after each page for progress reporting.
func extractLines(path string, onPage func(page, total int)) ([]string, error) {
	f, r, err := pdf.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	total := r.NumPage()
	var lines []string
	for i := 1; i <= total; i++ {
		if p := r.Page(i); !p.V.IsNull() {
			if rows, err := p.GetTextByRow(); err != nil {
				fmt.Fprintf(os.Stderr, "warning: page %d: %v\n", i, err)
			} else {
				for _, row := range rows {
					lines = append(lines, buildLine(row.Content))
				}
			}
		}
		if onPage != nil {
			onPage(i, total)
		}
	}
	return lines, nil
}

// buildLine joins the text fragments of one row, inserting a single space wherever
// there is a horizontal gap between fragments. This keeps multi-glyph tokens (SKU,
// the brand tag, model, qty, "EA") intact while separating columns, so the result
// tokenises reliably regardless of how the PDF chunked its text.
func buildLine(content pdf.TextHorizontal) string {
	sort.Sort(content) // by X
	var sb strings.Builder
	prevEnd := 0.0
	for i, t := range content {
		w := t.W
		if w <= 0 { // some PDFs omit width; approximate from glyph count
			w = float64(len([]rune(t.S))) * t.FontSize * 0.5
		}
		if i > 0 {
			thr := t.FontSize * 0.15
			if thr < 0.8 {
				thr = 0.8
			}
			if t.X-prevEnd > thr {
				sb.WriteByte(' ')
			}
		}
		sb.WriteString(t.S)
		prevEnd = t.X + w
	}
	return sb.String()
}
