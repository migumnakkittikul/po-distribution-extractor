# po-distribution-extractor

A small tool that takes a supplier purchase-order PDF and turns the
"distribute to branch" section into a single Excel sheet, one row per branch
line item.

This started as a manual job: export the PO, run it through an online
PDF-to-Excel converter, then clean up the mess by hand and look up each
branch's short name in a separate spreadsheet. It was slow, and easy to get
wrong, since a quantity landing in the wrong column quietly throws off the
totals. This does it in one step.

## What it does

- Reads the branch-distribution pages of the PO PDF and pulls out, per branch:
  line number, SKU, model, and quantity.
- Writes them into one worksheet (`ใบแบ่ง`) with a small per-model total table
  off to the side.
- Fills in each branch's Thai short name from a built-in list, so you don't
  have to look them up.
- Takes the quantity as the number printed right before the unit, which avoids
  the column-shift mistake that was easy to make by hand.

## Build

Needs Go 1.21 or newer.

```
go build -o po-distribution-extractor .
```

To cross-compile a Windows executable from any OS:

```
CGO_ENABLED=0 GOOS=windows GOARCH=amd64 go build -ldflags "-s -w -H windowsgui" -o po-distribution-extractor.exe .
```

The result is a single file with no runtime dependencies.

## Use

Double-click the program and a small menu opens:

- **Convert** picks a PO PDF, asks where to save, and writes the `.xlsx`.
- **Branches** shows the branch list and lets you add or rename one.

There's a command-line mode too, for scripting:

```
po-distribution-extractor -in purchase-order.pdf -out result.xlsx
```

## Branch names

The Thai branch names are built in, but you can add or change one without
rebuilding: drop a `branches.csv` next to the program with `code,thai_name`
rows. The Branches menu writes that file for you. It's read as UTF-8 (BOM or
not) and falls back to Windows-874, which is what Excel tends to save Thai CSVs
as.

## Sample

`testdata/sample-po.pdf` is a made-up purchase order. Run it:

```
po-distribution-extractor -in testdata/sample-po.pdf -out sample.xlsx
```

That produces one sheet with 12 branches and 23 line items, plus the per-model
totals (which add up to 85 units).

## Notes

Two things that were fiddlier than I expected:

- PDF text comes out as positioned fragments, not tidy rows. The columns are
  rebuilt from the horizontal gap between fragments, so a model code or a
  quantity never gets split apart or glued to the next column.
- The sheet keeps a large Thai font and sizes each column to its widest value,
  so the long branch names aren't cut off.

## License

MIT. See [LICENSE](LICENSE).
