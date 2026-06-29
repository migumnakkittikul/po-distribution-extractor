// Command po-distribution-extractor converts a purchase-order PDF into a
// per-branch distribution spreadsheet.
package main

import (
	"flag"
	"fmt"
	"os"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/ledongthuc/pdf"
)

const brandTag = "AQUARO"                  // the brand column value; model is the token right after it
const distributePrefix = "กระจายไปยังสาขา" // marks a "distribute to branch" block header

var (
	skuRe       = regexp.MustCompile(`^0\d{8}$`)        // SKU No. e.g. 060444193
	qtyRe       = regexp.MustCompile(`^\d+(?:\.\d+)?$`) // quantity e.g. 8.00
	poRe        = regexp.MustCompile(`ใบสั่งซื้อเลขที่\s*(\d+)`)
	branchHdrRe = regexp.MustCompile(`^(\d{4,6})\s+(.*\S)`) // "<code> <English name>"
)

type item struct {
	sku   int
	model string
	qty   int
}

type branch struct {
	code    string
	engName string
	items   []item
}

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
	branches, po := parse(lines)
	items := 0
	for _, b := range branches {
		items += len(b.items)
	}
	fmt.Printf("PO %s: %d branches, %d items\n", po, len(branches), items)
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

// parse walks the lines, tracking the current branch and collecting its items.
func parse(lines []string) ([]branch, string) {
	var branches []branch
	curIdx := -1 // index into branches; use index (not pointer) - append may reallocate
	poNumber := ""

	for _, line := range lines {
		if poNumber == "" {
			if m := poRe.FindStringSubmatch(line); m != nil {
				poNumber = m[1]
			}
		}

		// Branch block header, e.g. "กระจายไปยังสาขา 1001 Example Co. (Rangsit Branch)".
		if idx := strings.Index(line, distributePrefix); idx >= 0 {
			rest := strings.TrimSpace(line[idx+len(distributePrefix):])
			if m := branchHdrRe.FindStringSubmatch(rest); m != nil {
				branches = append(branches, branch{code: m[1], engName: strings.TrimSpace(m[2])})
				curIdx = len(branches) - 1
			}
			continue
		}

		if curIdx < 0 {
			continue // still in the PO section (pages 1-3); no branch yet
		}
		if it, ok := parseItem(line); ok {
			branches[curIdx].items = append(branches[curIdx].items, it)
		}
	}
	return branches, poNumber
}

// parseItem extracts (SKU, model, qty) from a distribution item row, or reports
// ok=false if the line is not an item row (header, footer, barcode-only, etc.).
func parseItem(line string) (item, bool) {
	toks := strings.Fields(line)
	eaIdx, skuIdx, vegIdx := -1, -1, -1
	for i, t := range toks {
		switch {
		case t == "EA":
			eaIdx = i // last occurrence = the unit column
		case t == brandTag:
			vegIdx = i // last occurrence = the brand column (model follows)
		case skuIdx == -1 && skuRe.MatchString(t):
			skuIdx = i // first 0XXXXXXXX = SKU No.
		}
	}
	if skuIdx == -1 || eaIdx <= 0 || vegIdx == -1 || vegIdx+1 >= len(toks) {
		return item{}, false
	}
	qtyTok := toks[eaIdx-1] // quantity sits immediately before "EA"
	if !qtyRe.MatchString(qtyTok) {
		return item{}, false
	}
	qty, err := strconv.ParseFloat(qtyTok, 64)
	if err != nil {
		return item{}, false
	}
	sku, err := strconv.Atoi(toks[skuIdx])
	if err != nil {
		return item{}, false
	}
	return item{sku: sku, model: toks[vegIdx+1], qty: int(qty + 0.5)}, true
}
