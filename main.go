// Command po-distribution-extractor converts a purchase-order PDF into a
// per-branch distribution spreadsheet.
package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/ledongthuc/pdf"
	"github.com/xuri/excelize/v2"
)

const sheetName = "ใบแบ่ง"

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
	outFlag := flag.String("out", "", "output .xlsx")
	invFlag := flag.String("invoice", "", "invoice number for cell E1 (optional)")
	flag.Parse()

	inPath, err := resolveInput(*inFlag, flag.Args())
	if err != nil {
		fmt.Fprintln(os.Stderr, "error:", err)
		os.Exit(1)
	}
	outPath := *outFlag
	if outPath == "" {
		outPath = strings.TrimSuffix(inPath, filepath.Ext(inPath)) + ".xlsx"
	}
	nb, ni, err := convert(inPath, outPath, *invFlag, func(int, string) {})
	if err != nil {
		fmt.Fprintln(os.Stderr, "error:", err)
		os.Exit(1)
	}
	fmt.Printf("Wrote %s: %d branches, %d items\n", outPath, nb, ni)
}

// convert runs the PDF -> sheet pipeline, reporting progress via the callback.
func convert(inPath, outPath, invoice string, progress func(pct int, msg string)) (branches, items int, err error) {
	lines, err := extractLines(inPath, func(page, total int) {
		pct := 5
		if total > 0 {
			pct = 5 + page*70/total
		}
		progress(pct, fmt.Sprintf("Reading PDF - page %d/%d...", page, total))
	})
	if err != nil {
		return 0, 0, err
	}
	progress(82, "Finding branches...")
	bs, poNumber := parse(lines)
	if len(bs) == 0 {
		return 0, 0, fmt.Errorf("no branch blocks found - is this a Flow Through purchase-order PDF?")
	}
	if poNumber == "" {
		base := filepath.Base(inPath)
		poNumber = strings.TrimSuffix(base, filepath.Ext(base))
	}
	progress(92, "Writing Excel...")
	if err := writeXLSX(outPath, bs, poNumber, invoice); err != nil {
		return 0, 0, err
	}
	for _, b := range bs {
		items += len(b.items)
	}
	progress(100, "Done")
	return len(bs), items, nil
}

func resolveInput(inFlag string, args []string) (string, error) {
	cand := inFlag
	if cand == "" && len(args) > 0 {
		cand = args[0]
	}
	if cand == "" {
		return "", fmt.Errorf("no input PDF given")
	}
	if !fileExists(cand) {
		return "", fmt.Errorf("file not found: %s", cand)
	}
	return cand, nil
}

func fileExists(p string) bool {
	info, err := os.Stat(p)
	return err == nil && !info.IsDir()
}

// extractLines returns one reconstructed text line per visual row of the PDF.
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
// there is a horizontal gap between fragments.
func buildLine(content pdf.TextHorizontal) string {
	sort.Sort(content) // by X
	var sb strings.Builder
	prevEnd := 0.0
	for i, t := range content {
		w := t.W
		if w <= 0 {
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
	curIdx := -1
	poNumber := ""

	for _, line := range lines {
		if poNumber == "" {
			if m := poRe.FindStringSubmatch(line); m != nil {
				poNumber = m[1]
			}
		}
		if idx := strings.Index(line, distributePrefix); idx >= 0 {
			rest := strings.TrimSpace(line[idx+len(distributePrefix):])
			if m := branchHdrRe.FindStringSubmatch(rest); m != nil {
				branches = append(branches, branch{code: m[1], engName: strings.TrimSpace(m[2])})
				curIdx = len(branches) - 1
			}
			continue
		}
		if curIdx < 0 {
			continue
		}
		if it, ok := parseItem(line); ok {
			branches[curIdx].items = append(branches[curIdx].items, it)
		}
	}
	return branches, poNumber
}

func parseItem(line string) (item, bool) {
	toks := strings.Fields(line)
	eaIdx, skuIdx, vegIdx := -1, -1, -1
	for i, t := range toks {
		switch {
		case t == "EA":
			eaIdx = i
		case t == brandTag:
			vegIdx = i
		case skuIdx == -1 && skuRe.MatchString(t):
			skuIdx = i
		}
	}
	if skuIdx == -1 || eaIdx <= 0 || vegIdx == -1 || vegIdx+1 >= len(toks) {
		return item{}, false
	}
	qtyTok := toks[eaIdx-1]
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

// writeXLSX writes the branches into the ใบแบ่ง sheet, one row per line item.
func writeXLSX(path string, branches []branch, poNumber, invoice string) error {
	f := excelize.NewFile()
	defer f.Close()

	idx, err := f.NewSheet(sheetName)
	if err != nil {
		return err
	}
	f.SetActiveSheet(idx)
	if err := f.DeleteSheet("Sheet1"); err != nil {
		return err
	}

	var firstErr error
	set := func(col, row int, v interface{}) {
		if firstErr != nil {
			return
		}
		cell, err := excelize.CoordinatesToCellName(col, row)
		if err != nil {
			firstErr = err
			return
		}
		if err := f.SetCellValue(sheetName, cell, v); err != nil {
			firstErr = err
		}
	}

	set(1, 1, "ใบสั่งซื้อเลขที่ "+poNumber)
	if invoice != "" {
		set(5, 1, invoice)
	}
	for i, h := range []string{"ลำดับ", "SKU No.", "รุ่น", "จำนวน", "สาขา"} {
		set(i+1, 2, h)
	}

	row := 3
	for _, b := range branches {
		set(2, row, b.code+" "+b.engName)
		row++
		for i, it := range b.items {
			set(1, row, i+1)
			set(2, row, it.sku)
			set(3, row, it.model)
			set(4, row, it.qty)
			row++
		}
	}
	if firstErr != nil {
		return firstErr
	}
	return f.SaveAs(path)
}
