// Command po-distribution-extractor converts a "Flow Through" purchase-order
// PDF (e.g. sample-po.pdf) into the ใบแบ่ง (branch-distribution) sheet of an .xlsx.
//
// It replaces the old manual pipeline (online PDF->Excel tool -> Sheet1 -> manual
// cleanup + cross-reference). Each ใบแบ่ง component is produced as follows:
//
//	A1            "ใบสั่งซื้อเลขที่ <PO>"   <- PDF  (regex "ใบสั่งซื้อเลขที่ (\d+)", else filename)
//	E1            (left blank)             <- not in any source; set only via optional -invoice flag
//	row 2         column headers           <- static
//	branch row B  "<code> <English name>"  <- PDF  ("กระจายไปยังสาขา ..." block header)
//	branch row E  Thai short name          <- embedded branch map (+ optional branches.csv override)
//	item   A      line number              <- per-branch counter (1,2,3,...)
//	item   B      SKU                       <- PDF  ("SKU No." 0XXXXXXXX -> int, drops leading 0)
//	item   C      model (รุ่น)              <- PDF  (token after the last brand tag on the row)
//	item   D      quantity (จำนวน)          <- PDF  (the number immediately before "EA")
//	G3:H..        model -> total qty       <- computed (SUM of qty grouped by model)
package main

import (
	"bufio"
	"bytes"
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"unicode/utf8"

	"github.com/ledongthuc/pdf"
	"github.com/ncruces/zenity"
	"github.com/xuri/excelize/v2"
	"golang.org/x/text/encoding/charmap"
)

// branchThaiName is the built-in branch code -> Thai short-name list (column
// E on each branch header row). An optional branches.csv next to the program
// adds or overrides entries without needing a rebuild.
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

// sheetName is the output tab name. The original workbook's tab is "ใบแบ่ง " with a
// trailing space; we use the clean form here. Change if you need a byte-exact match.
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

func main() {
	inFlag := flag.String("in", "", "input PDF (headless mode); if omitted, a file picker opens")
	outFlag := flag.String("out", "", "output .xlsx (headless mode)")
	invFlag := flag.String("invoice", "", "invoice number for cell E1 (optional; not present in the PDF)")
	flag.Parse()

	loadBranchCSV() // optional branches.csv (next to the .exe) can add/override names

	// No input on the command line (i.e. double-clicked) -> graphical mode: a small
	// menu to convert a PO or manage the branch list. Otherwise run headless
	// (used for automation and testing).
	if *inFlag == "" && len(flag.Args()) == 0 {
		for {
			switch guiMenu() {
			case "convert":
				runConvertGUI(*invFlag)
				return
			case "manage":
				manageBranchesGUI()
			default:
				return
			}
		}
	}

	inPath, err := resolveInput(*inFlag, flag.Args())
	if err != nil {
		fail(err)
	}
	outPath := *outFlag
	if outPath == "" {
		outPath = strings.TrimSuffix(inPath, filepath.Ext(inPath)) + ".xlsx"
	}
	fmt.Printf("Reading: %s\n", inPath)
	nb, ni, err := convert(inPath, outPath, *invFlag, func(int, string) {}) // headless: no progress UI
	if err != nil {
		fail(err)
	}
	fmt.Printf("\nDone.  Wrote %s\n  sheet %q - %d branches - %d items\n", outPath, sheetName, nb, ni)
	pause()
}

// convert runs the full PDF-to-ใบแบ่ง pipeline, reporting real progress via the
// callback. PDF reading is ~99% of the work, so the bar tracks it page-by-page.
func convert(inPath, outPath, invoice string, progress func(pct int, msg string)) (branches, items int, err error) {
	lines, err := extractLines(inPath, func(page, total int) {
		pct := 5
		if total > 0 {
			pct = 5 + page*70/total // reading spans 5%..75%
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
	if poNumber == "" { // fall back to the input filename without extension
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

// guiMenu shows the start screen and returns "convert", "manage", or "exit".
func guiMenu() string {
	// Button labels are kept short - zenity's Windows buttons are a fixed width
	// and clip long text - so the descriptions live in the message body instead.
	err := zenity.Question(
		"Purchase order to distribution sheet.\n\nWhat would you like to do?\n\n"+
			"Convert   -   build the Excel file from a PO PDF\n"+
			"Branches   -   view or add branch names",
		zenity.Title("PO Distribution"),
		zenity.OKLabel("Convert"),
		zenity.ExtraButton("Branches"),
		zenity.CancelLabel("Exit"),
	)
	switch {
	case err == nil:
		return "convert"
	case errors.Is(err, zenity.ErrExtraButton):
		return "manage"
	default:
		return "exit"
	}
}

// manageBranchesGUI lets the user view all known branches and add/update names.
// Additions are saved to branches.csv next to the program.
func manageBranchesGUI() {
	for {
		codes := sortedBranchCodes()
		items := make([]string, len(codes))
		for i, c := range codes {
			items[i] = c + "    " + branchThaiName[c]
		}
		_, err := zenity.List(
			fmt.Sprintf("%d branches (code -> Thai name).\nClick “Add” to add a new branch - or to rename one, type its existing code.", len(codes)),
			items,
			zenity.Title("Branches"),
			zenity.OKLabel("Add"),
			zenity.CancelLabel("Close"),
			zenity.Height(460),
		)
		if errors.Is(err, zenity.ErrCanceled) || (err != nil && !errors.Is(err, zenity.ErrExtraButton)) {
			return // Close (or dialog unavailable)
		}
		addBranchGUI() // OK ("Add / edit") pressed
	}
}

// addBranchGUI prompts for a code + Thai name and persists it to branches.csv.
func addBranchGUI() {
	code, err := zenity.Entry("Branch code (digits only, e.g. 61605):", zenity.Title("Add / edit a branch"))
	if err != nil {
		return // canceled
	}
	code = strings.TrimSpace(code)
	if !digitsRe.MatchString(code) {
		_ = zenity.Error("The branch code must be digits only (e.g. 61605).", zenity.Title("Invalid code"))
		return
	}
	existing := branchThaiName[code]
	name, err := zenity.Entry(
		fmt.Sprintf("Thai name for branch %s:", code),
		zenity.Title("Add / edit a branch"),
		zenity.EntryText(existing),
	)
	if err != nil {
		return
	}
	name = strings.TrimSpace(name)
	if name == "" {
		_ = zenity.Error("The name cannot be empty.", zenity.Title("Invalid name"))
		return
	}
	if err := saveBranchTo(branchCSVPath(), code, name); err != nil {
		_ = zenity.Error("Could not save branches.csv:\n"+err.Error(), zenity.Title("Save failed"))
		return
	}
	branchThaiName[code] = name // reflect immediately in the list
	verb := "Added"
	if existing != "" {
		verb = "Updated"
	}
	_ = zenity.Info(
		fmt.Sprintf("%s branch %s -> %s\n\nSaved to:\n%s", verb, code, name, branchCSVPath()),
		zenity.Title("Saved"),
	)
}

func sortedBranchCodes() []string {
	codes := make([]string, 0, len(branchThaiName))
	for c := range branchThaiName {
		codes = append(codes, c)
	}
	sort.Slice(codes, func(i, j int) bool {
		a, _ := strconv.Atoi(codes[i])
		b, _ := strconv.Atoi(codes[j])
		return a < b
	})
	return codes
}

// runConvertGUI is the convert flow: choose the PDF, choose where to save, show a
// progress window, then a success/error message - all via native dialogs.
func runConvertGUI(invoice string) {
	inPath, err := zenity.SelectFile(
		zenity.Title("Choose the purchase-order PDF"),
		zenity.FileFilters{{Name: "PDF files", Patterns: []string{"*.pdf"}}},
	)
	if errors.Is(err, zenity.ErrCanceled) {
		return
	}
	if err != nil {
		zenity.Error(err.Error(), zenity.Title("Error"))
		return
	}

	defPath := filepath.Join(filepath.Dir(inPath),
		strings.TrimSuffix(filepath.Base(inPath), filepath.Ext(inPath))+".xlsx")
	outPath, err := zenity.SelectFileSave(
		zenity.Title("Save the Excel file as"),
		zenity.ConfirmOverwrite(),
		zenity.Filename(defPath),
		zenity.FileFilters{{Name: "Excel files", Patterns: []string{"*.xlsx"}}},
	)
	if errors.Is(err, zenity.ErrCanceled) {
		return
	}
	if err != nil {
		zenity.Error(err.Error(), zenity.Title("Error"))
		return
	}
	if filepath.Ext(outPath) == "" {
		outPath += ".xlsx"
	}

	dlg, derr := zenity.Progress(
		zenity.Title("PO Distribution"),
		zenity.MaxValue(100),
		zenity.NoCancel(),
	)
	progress := func(pct int, msg string) {
		if derr == nil {
			dlg.Value(pct)
			dlg.Text(msg)
		}
	}

	nb, ni, cerr := convert(inPath, outPath, invoice, progress)
	if derr == nil {
		dlg.Complete()
		dlg.Close()
	}
	if cerr != nil {
		zenity.Error(cerr.Error(), zenity.Title("Conversion failed"))
		return
	}
	zenity.Info(
		fmt.Sprintf("Done!\n\nSaved to:\n%s\n\n%d branches - %d items", outPath, nb, ni),
		zenity.Title("PO Distribution"),
	)
}

// resolveInput picks the input PDF for headless mode: the -in flag, else the first
// positional argument. (Double-click / no-argument launches use the GUI picker.)
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

var stdin = bufio.NewReader(os.Stdin)

var digitsRe = regexp.MustCompile(`^\d+$`)

// loadBranchCSV looks for an optional "branches.csv" next to the program (or in the
// current folder) and merges "code,thai_name" rows into branchThaiName. A row whose
// code already exists overrides the embedded name; a new code is added. It is
// optional: if no file is found, the embedded 122-branch table is used unchanged.
// This lets new branches be added without recompiling - just edit the CSV.
func loadBranchCSV() {
	for _, dir := range csvSearchDirs() {
		p := filepath.Join(dir, "branches.csv")
		data, err := os.ReadFile(p)
		if err != nil {
			continue
		}
		n := mergeBranchCSV(data)
		fmt.Printf("Loaded %d branch name(s) from %s\n", n, p)
		return
	}
}

func csvSearchDirs() []string {
	var dirs []string
	if exe, err := os.Executable(); err == nil {
		dirs = append(dirs, filepath.Dir(exe))
	}
	if cwd, err := os.Getwd(); err == nil {
		dirs = append(dirs, cwd)
	}
	return dirs
}

func mergeBranchCSV(data []byte) int {
	r := csv.NewReader(strings.NewReader(decodeText(data)))
	r.FieldsPerRecord = -1 // allow ragged rows
	r.TrimLeadingSpace = true
	n := 0
	for {
		rec, err := r.Read()
		if err != nil {
			break
		}
		if len(rec) < 2 {
			continue
		}
		code := strings.TrimSpace(rec[0])
		if !digitsRe.MatchString(code) { // skips a header row, comments, blanks
			continue
		}
		branchThaiName[code] = strings.TrimSpace(rec[1])
		n++
	}
	return n
}

// decodeText returns UTF-8 text from CSV bytes, tolerating a UTF-8 BOM and falling
// back to Windows-874 (Thai) - the encoding Excel often uses when saving Thai CSVs.
func decodeText(b []byte) string {
	b = bytes.TrimPrefix(b, []byte{0xEF, 0xBB, 0xBF})
	if utf8.Valid(b) {
		return string(b)
	}
	if out, err := charmap.Windows874.NewDecoder().Bytes(b); err == nil {
		return string(out)
	}
	return string(b)
}

// branchCSVPath is the branches.csv location: next to the program executable.
func branchCSVPath() string {
	if exe, err := os.Executable(); err == nil {
		return filepath.Join(filepath.Dir(exe), "branches.csv")
	}
	return "branches.csv"
}

// saveBranchTo adds or updates one branch in the CSV at path, preserving any
// existing rows. The file is (re)written as UTF-8 with a BOM (so Excel reads Thai
// correctly), with a header and rows sorted by code.
func saveBranchTo(path, code, name string) error {
	entries := map[string]string{}
	if data, err := os.ReadFile(path); err == nil {
		r := csv.NewReader(strings.NewReader(decodeText(data)))
		r.FieldsPerRecord = -1
		r.TrimLeadingSpace = true
		for {
			rec, e := r.Read()
			if e != nil {
				break
			}
			if len(rec) >= 2 {
				if c := strings.TrimSpace(rec[0]); digitsRe.MatchString(c) {
					entries[c] = strings.TrimSpace(rec[1])
				}
			}
		}
	}
	entries[code] = name

	codes := make([]string, 0, len(entries))
	for c := range entries {
		codes = append(codes, c)
	}
	sort.Slice(codes, func(i, j int) bool {
		a, _ := strconv.Atoi(codes[i])
		b, _ := strconv.Atoi(codes[j])
		return a < b
	})

	var sb strings.Builder
	sb.WriteString("code,thai_name\r\n") //  = UTF-8 BOM so Excel detects UTF-8
	for _, c := range codes {
		n := entries[c]
		if strings.ContainsAny(n, ",\"\n\r") {
			n = `"` + strings.ReplaceAll(n, `"`, `""`) + `"`
		}
		sb.WriteString(c + "," + n + "\r\n")
	}
	out := append([]byte{0xEF, 0xBB, 0xBF}, sb.String()...) // UTF-8 BOM so Excel detects UTF-8
	return os.WriteFile(path, out, 0o644)
}

func pause() {
	fmt.Print("\nPress Enter to close...")
	stdin.ReadString('\n')
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
	sku, err := strconv.Atoi(toks[skuIdx]) // "060444193" -> 60444193 (base 10, drops leading 0)
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

func fail(err error) {
	fmt.Fprintln(os.Stderr, "\nError:", err)
	pause()
	os.Exit(1)
}
