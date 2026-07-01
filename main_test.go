package main

import (
	"bytes"
	"encoding/csv"
	"os"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

const samplePDF = "testdata/sample-po.pdf"

func mustLines(tb testing.TB) []string {
	lines, err := extractLines(samplePDF, nil)
	if err != nil {
		tb.Fatalf("extractLines: %v", err)
	}
	return lines
}

func TestParseCounts(t *testing.T) {
	bs, po := parse(mustLines(t))
	if po != "1042000789" {
		t.Errorf("PO = %q, want 1042000789", po)
	}
	if len(bs) != 12 {
		t.Errorf("branches = %d, want 12", len(bs))
	}
	items := 0
	for _, b := range bs {
		items += len(b.items)
	}
	if items != 23 {
		t.Errorf("items = %d, want 23", items)
	}
}

// A quantity is the number right before "EA"; check one branch's rows parse.
func TestBranchQuantities(t *testing.T) {
	bs, _ := parse(mustLines(t))
	var rangsit *branch
	for i := range bs {
		if bs[i].code == "1001" {
			rangsit = &bs[i]
		}
	}
	if rangsit == nil {
		t.Fatal("branch 1001 not found")
	}
	got := map[string]int{}
	for _, it := range rangsit.items {
		got[it.model] = it.qty
	}
	if got["A-100"] != 3 || got["B-205"] != 5 {
		t.Errorf("branch 1001 = %v, want A-100:3 B-205:5", got)
	}
}

func TestModelTotals(t *testing.T) {
	bs, _ := parse(mustLines(t))
	tot := map[string]int{}
	grand := 0
	for _, b := range bs {
		for _, it := range b.items {
			tot[it.model] += it.qty
			grand += it.qty
		}
	}
	want := map[string]int{"A-100": 11, "B-205": 16, "S-300": 10, "T-410": 19, "H-520": 14, "K-630": 15}
	for m, w := range want {
		if tot[m] != w {
			t.Errorf("%s total = %d, want %d", m, tot[m], w)
		}
	}
	if grand != 85 {
		t.Errorf("grand total = %d, want 85", grand)
	}
}

func TestEndToEnd(t *testing.T) {
	out := t.TempDir() + "/out.xlsx"
	nb, ni, err := convert(samplePDF, out, "", func(int, string) {})
	if err != nil {
		t.Fatal(err)
	}
	if nb != 12 || ni != 23 {
		t.Fatalf("convert returned %d branches, %d items; want 12, 23", nb, ni)
	}
	f, err := excelize.OpenFile(out)
	if err != nil {
		t.Fatalf("reopen: %v", err)
	}
	defer f.Close()
	if got := f.GetSheetList(); len(got) != 1 || got[0] != sheetName {
		t.Errorf("sheets = %v, want [%q]", got, sheetName)
	}
	if v, _ := f.GetCellValue(sheetName, "A1"); v != "ใบสั่งซื้อเลขที่ 1042000789" {
		t.Errorf("A1 = %q", v)
	}
}

// saveBranchTo backs the in-app "add/edit branch": BOM, Thai, comma-quoting,
// update-in-place, preserve existing, sort by code.
func TestSaveBranchTo(t *testing.T) {
	p := t.TempDir() + "/branches.csv"
	for _, e := range []struct{ code, name string }{
		{"2999", "ทดสอบเก่า"},
		{"1001", "เทส"},
		{"2888", "ก,ข"},
		{"2999", "ทดสอบใหม่"},
	} {
		if err := saveBranchTo(p, e.code, e.name); err != nil {
			t.Fatalf("saveBranchTo: %v", err)
		}
	}
	data, err := os.ReadFile(p)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.HasPrefix(data, []byte{0xEF, 0xBB, 0xBF}) {
		t.Error("missing UTF-8 BOM")
	}
	got := map[string]string{}
	r := csv.NewReader(strings.NewReader(decodeText(data)))
	r.FieldsPerRecord = -1
	for {
		rec, e := r.Read()
		if e != nil {
			break
		}
		if len(rec) >= 2 && digitsRe.MatchString(strings.TrimSpace(rec[0])) {
			got[strings.TrimSpace(rec[0])] = rec[1]
		}
	}
	if got["1001"] != "เทส" {
		t.Errorf("1001 = %q, want เทส", got["1001"])
	}
	if got["2888"] != "ก,ข" {
		t.Errorf("2888 = %q, want ก,ข (comma round-trip)", got["2888"])
	}
	if got["2999"] != "ทดสอบใหม่" {
		t.Errorf("2999 = %q, want ทดสอบใหม่ (updated)", got["2999"])
	}
	dec := decodeText(data)
	if strings.Index(dec, "1001") > strings.Index(dec, "2999") {
		t.Error("rows not sorted by code")
	}
}
