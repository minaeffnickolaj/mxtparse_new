package main

/* GNU GPL v2.0
	This utilite was written for parsing
	Rigla drugstores RDP connections from XLSX
	to .mxtsessions (MobaXTerm) file format.
	This version use goroutines and channels for
	best performance with large files.
   Author: Minaev N.
*/

import (
	"flag"
	"fmt"
	"runtime"
	"strconv"
	"text/template"

	"github.com/xuri/excelize/v2"
)

// flag variables initialization, pointer to *type of flag
var (
	InputFile     *string
	OutputFile    *string
	FullFile      *bool
	DisplayOutput *bool
)

// flag declarations according to types of variables
func init() {
	InputFile = flag.String("InputFile", "apt.xlsx", "Filepath to input XLSX file")
	OutputFile = flag.String("OutputFile", "apt.mxtsessions", "Filepath to output .mxt file")
	FullFile = flag.Bool("FullFile", true, "Add cashiers PC's to file?")
	DisplayOutput = flag.Bool("ShowOutput", false, "Output parsed connections to terminal?")
}
func main() {
	flag.Parse()      // get flags
	PrintParameters() // show parameters of task
	XLSX := OpenXLSXFile(*InputFile)
	XLSXChannelOut := make(chan RDP)             // channel to XLSX out
	ReadXLSXFile(XLSX, "Аптеки", XLSXChannelOut) // read from file to channel
	if !*FullFile {
		// declare template
		MXTTemplate := template.New("MobaXTermTemplate")
		MXTTemplateText := "\n\n[Bookmarks_{{.RecordNum}}]\nSubRep={{.RKName}}\\{{.AptName}}\nImgNum=41\n{{.AptName}}({{.APCode}})=#91#4%{{.ServerAddress}}.apt.rigla.ru%10433%[{{.Username}}]%0%-1%-1%-1%-1%0%0%-1%%%%%0%0%%-1%%-1%-1%0%-1%0%-1#MobaFont%10%0%0%-1%15%236,236,236%30,30,30%180,180,192%0%-1%0%%xterm%-1%-1%_Std_Colors_0_%80%24%0%1%-1%<none>%%0%0%-1#0# #-1"
		MXTTemplate.Parse(MXTTemplateText)
		for i := 0; i <= runtime.NumCPU(); i++ { //goroutine count will be equal CPU core count
			//	go ParseTemplate()
		}
	} else {
		for i := 0; i <= runtime.NumCPU(); i++ {

		}
	}
}

// read struct from *RDP channel to text/template lib, result to chan []byte to WriteFile()
func ParseTemplate(TemplateText *template.Template, IncomeRDPChannel chan RDP, ChannelOut chan []byte) {
	// get *RDP from chan
	RDParsed := <-IncomeRDPChannel
	//TemplateText.Execute(
}

func WriteFile() {

}

// read file line by line from XLSX list, creating RDP struct and take it to free goroutine
func ReadXLSXFile(XLSXFile *excelize.File, ListName string, ChannelOut chan RDP) {
	Rows, err := XLSXFile.GetRows(ListName)
	if err != nil {
		println(err)
		return
	}
	for i, row := range Rows {
		/*
			Excel row in Excelize realized like [][]string, when each
			cell is [i][y] - i is Row, y is Cell
		*/
		MXTConnection := new(RDP) //create *RDP struct
		MXTConnection.RecordNum = strconv.Itoa(i + 1)
		MXTConnection.APCode = row[1]
		MXTConnection.RKName = row[2]
		MXTConnection.AptName = row[0]
		MXTConnection.ServerAddress = row[3]
		MXTConnection.Username = "efarma"
		ChannelOut <- *MXTConnection
	}
}

func PrintParameters() { //preview parameters before continue
	println("XLSX to MobaXTerm parser v0.1 \n")
	fmt.Println("Path to XLSX: ", *InputFile)
	fmt.Println("Path to .mxtsessions: ", *OutputFile)
	if !*FullFile {
		println("File will be parsed without VNC connections to cashiers PC's")
	} else {
		println("File will be parsed with VNC connections to cashiers PC's")
	}
	if *DisplayOutput {
		println("Progress will be shown on terminal output \n")
	} else {
		println("Silent mode activated \n")
	}
}

func OpenXLSXFile(Filepath string) *excelize.File {
	File, Error := excelize.OpenFile(Filepath)
	if Error != nil {
		fmt.Println("File can't be opened", Error)
		return nil
	}
	defer func() {
		if Error := File.Close(); Error != nil {
			fmt.Println(Error)
		}
	}()
	return File
}
