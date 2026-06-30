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

// Output styling, chosen to match the original ใบแบ่ง sheet (a large Thai font).
const (
	fontName = "Angsana New"
	fontSize = 20.0
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

var branchThaiName = map[string]string{
	"1001": "รังสิต",
	"1002": "บางนา",
	"1003": "นนทบุรี",
	"1004": "ลาดพร้าว",
	"1005": "บางแค",
	"1006": "มีนบุรี",
	"1007": "ปทุมธานี",
	"1008": "สมุทรปราการ",
	"1009": "ชลบุรี",
	"1010": "ระยอง",
	"1011": "เชียงใหม่",
	"1012": "เชียงราย",
	"1013": "ขอนแก่น",
	"1014": "อุดรธานี",
	"1015": "นครราชสีมา",
	"1016": "อุบลราชธานี",
	"1017": "ภูเก็ต",
	"1018": "สุราษฎร์ธานี",
	"1019": "สงขลา",
	"1020": "หาดใหญ่",
	"1021": "พิษณุโลก",
	"1022": "นครสวรรค์",
	"1023": "ราชบุรี",
	"1024": "กาญจนบุรี",
	"1025": "ลพบุรี",
	"1026": "อยุธยา",
	"1027": "สระบุรี",
	"1028": "ฉะเชิงเทรา",
	"1029": "นครปฐม",
	"1030": "เพชรบุรี",
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

// writeXLSX renders the branches into the ใบแบ่ง sheet of a fresh workbook.
func writeXLSX(path string, branches []branch, poNumber, invoice string) error {
	f := excelize.NewFile()
	defer f.Close()

	idx, err := f.NewSheet(sheetName)
	if err != nil {
		return err
	}
	f.SetActiveSheet(idx)
	if err := f.DeleteSheet("Sheet1"); err != nil { // drop the default sheet
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

	// Title row + invoice number.
	set(1, 1, "ใบสั่งซื้อเลขที่ "+poNumber) // A1
	if invoice != "" {
		set(5, 1, invoice) // E1
	}
	// Column headers (row 2).
	for i, h := range []string{"ลำดับ", "SKU No.", "รุ่น", "จำนวน", "สาขา"} {
		set(i+1, 2, h)
	}

	// Branch blocks + item rows.
	row := 3
	totals := map[string]int{}
	for _, b := range branches {
		set(2, row, b.code+" "+b.engName) // B: code + English name (from PDF)
		if name, ok := branchThaiName[b.code]; ok {
			set(5, row, name) // E: Thai short name (cross-reference)
		} else {
			fmt.Fprintf(os.Stderr, "warning: no Thai name for branch %s (%s); E left blank\n", b.code, b.engName)
		}
		row++
		for i, it := range b.items {
			set(1, row, i+1)      // A: line number
			set(2, row, it.sku)   // B: SKU
			set(3, row, it.model) // C: model
			set(4, row, it.qty)   // D: quantity
			totals[it.model] += it.qty
			row++
		}
	}

	// Summary table (รุ่น -> total qty) in G:H, models sorted ascending.
	set(7, 3, "รุ่น")
	set(8, 3, "จำนวน")
	models := make([]string, 0, len(totals))
	for m := range totals {
		models = append(models, m)
	}
	sort.Strings(models)
	for i, m := range models {
		set(7, 4+i, m)
		set(8, 4+i, totals[m])
	}
	if firstErr != nil {
		return firstErr
	}

	// ---- formatting to match the original ใบแบ่ง sheet ----
	lastRow := row - 1         // last written data row
	sumLast := 3 + len(models) // last summary row (G/H)

	mk := func(bold bool, h, v string) int {
		id, e := f.NewStyle(&excelize.Style{
			Font:      &excelize.Font{Family: fontName, Size: fontSize, Bold: bold},
			Alignment: &excelize.Alignment{Horizontal: h, Vertical: v},
		})
		if e != nil && firstErr == nil {
			firstErr = e
		}
		return id
	}
	base := mk(false, "", "")    // plain Angsana 20 (default for the sheet)
	title := mk(true, "", "top") // row 1
	hdrC := mk(true, "center", "top")
	hdrL := mk(true, "left", "top")
	dCenter := mk(false, "center", "top")  // line#, model, qty
	dLeft := mk(false, "left", "top")      // SKU / branch name
	dEmid := mk(false, "center", "center") // branch Thai name (col E)
	sumData := mk(false, "left", "")
	if firstErr != nil {
		return firstErr
	}

	// Whole-sheet default font (also covers empty cells).
	_ = f.SetColStyle(sheetName, "A:H", base)

	// Auto-size columns to their widest content so nothing is cut off - especially
	// the long branch names in column B. Width units are characters of the default
	// font; the cells use a larger font, so we scale up and pad generously.
	rlen := func(s string) int { return len([]rune(s)) }
	maxA, maxB, maxC, maxD, maxE := rlen("ลำดับ"), rlen("SKU No."), rlen("รุ่น"), rlen("จำนวน"), rlen("สาขา")
	maxG, maxH := rlen("รุ่น"), rlen("จำนวน")
	up := func(p *int, n int) {
		if n > *p {
			*p = n
		}
	}
	for _, b := range branches {
		up(&maxB, rlen(b.code+" "+b.engName))
		up(&maxE, rlen(branchThaiName[b.code]))
		up(&maxA, len(strconv.Itoa(len(b.items))))
		for _, it := range b.items {
			up(&maxB, len(strconv.Itoa(it.sku)))
			up(&maxC, rlen(it.model))
			up(&maxD, len(strconv.Itoa(it.qty)))
		}
	}
	for _, m := range models {
		up(&maxG, rlen(m))
		up(&maxH, len(strconv.Itoa(totals[m])))
	}
	width := func(n int) float64 {
		w := float64(n)*1.5 + 2 // scale generously for the large Angsana font + padding
		if w < 5 {
			w = 5
		} else if w > 100 {
			w = 100
		}
		return w
	}
	for _, cw := range []struct {
		c string
		w float64
	}{{"A", width(maxA)}, {"B", width(maxB)}, {"C", width(maxC)}, {"D", width(maxD)},
		{"E", width(maxE)}, {"F", 2.625}, {"G", width(maxG)}, {"H", width(maxH)}} {
		_ = f.SetColWidth(sheetName, cw.c, cw.c, cw.w)
	}
	rh, custom := 29.25, true
	_ = f.SetSheetProps(sheetName, &excelize.SheetPropsOptions{DefaultRowHeight: &rh, CustomHeight: &custom})

	// Title (row 1) and bold column headers (row 2).
	_ = f.SetCellStyle(sheetName, "A1", "H1", title)
	_ = f.SetCellStyle(sheetName, "A2", "E2", hdrC)
	_ = f.SetCellStyle(sheetName, "B2", "B2", hdrL)

	// Data columns (rows 3..last): line#/model/qty centered, SKU/name left, branch centred.
	_ = f.SetCellStyle(sheetName, "A3", fmt.Sprintf("A%d", lastRow), dCenter)
	_ = f.SetCellStyle(sheetName, "B3", fmt.Sprintf("B%d", lastRow), dLeft)
	_ = f.SetCellStyle(sheetName, "C3", fmt.Sprintf("D%d", lastRow), dCenter)
	_ = f.SetCellStyle(sheetName, "E3", fmt.Sprintf("E%d", lastRow), dEmid)

	// Summary table (G:H).
	_ = f.SetCellStyle(sheetName, "G3", "G3", hdrC)
	_ = f.SetCellStyle(sheetName, "H3", "H3", hdrL)
	_ = f.SetCellStyle(sheetName, "G4", fmt.Sprintf("H%d", sumLast), sumData)

	// Header-row dropdowns + repeat the header row when printing (as in the original).
	_ = f.AutoFilter(sheetName, fmt.Sprintf("A2:E%d", lastRow), nil)
	_ = f.SetDefinedName(&excelize.DefinedName{Name: "_xlnm.Print_Titles", RefersTo: "'" + sheetName + "'!$2:$2", Scope: sheetName})

	return f.SaveAs(path)
}
