// Command po-distribution-extractor converts a purchase-order PDF into a
// per-branch distribution spreadsheet.
package main

import (
	"flag"
	"fmt"
)

func main() {
	inFlag := flag.String("in", "", "input PDF")
	outFlag := flag.String("out", "", "output .xlsx")
	flag.Parse()
	_ = outFlag

	if *inFlag == "" {
		fmt.Println("usage: po-distribution-extractor -in <file.pdf> -out <file.xlsx>")
		return
	}
	fmt.Println("not implemented yet")
}
