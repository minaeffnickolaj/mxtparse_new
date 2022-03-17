package mxtparsenew

/* GNU GPL v2.0
	This utilite was written for parsing
	Rigla drugstores RDP connections from XLSX
	to .mxtsessions (MobaXTerm) file format.
	This version use goroutines and pipes for
	best performance with large files.
   Author: Minaev N.
*/

import (
	"flag"
	_ "fmt"
)

// declation flags
var InputFilePath = flag.String("input", "apt.xlsx", "Path to XLSX file")
var OutputFilePath = flag.String("output", "list.mxtsessions", "Path to .mxtsessions file")

// bool variable for short or full list of connections (include cashier's PC or not)
var ParseCashiersPCToFile bool

// full or short version file?
func init() {
	flag.BoolVar(&ParseCashiersPCToFile, "Add cashiers PC to file?", false, "Include cashiers PC's")
}

func mxtparsenew() {
	//fmt.Print()
}
