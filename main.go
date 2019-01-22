package main

import (
	"bufio"
	"bytes"
	"encoding/xml"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"golang.org/x/text/encoding"
	"golang.org/x/text/encoding/charmap"
	"gopkg.in/cheggaaa/pb.v1"
	"io"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strconv"
)

const HELLO = "YML to Excel converter"

func main() {

	dir, err := os.Getwd()

	if err != nil {
		fmt.Println(err.Error())
		fmt.Println("cant find myself")
		fmt.Println("press ENTER")
		fmt.Scanln()
	}
	dir = dir + "\\"
	//dir = filepath.Dir("C:\\Users\\bolshakov\\go\\src\\ymltoxlsx\\")

	processDir(dir)

	fmt.Println("press ENTER")
	fmt.Scanln()
}

func askForGetParams() bool {
	fmt.Println("Do you need params page?")
	fmt.Println("1 - YES")
	fmt.Println("2 - NO")
	reader := bufio.NewReader(os.Stdin)
	r, _, err := reader.ReadRune()
	if err != nil {
		panic(err)
	}
	if r == bytes.Runes([]byte("1"))[0] {
		return true
	}
	return false
}

func askForXlsx() bool {
	fmt.Println("Do you need XLSX?")
	fmt.Println("1 - YES")
	fmt.Println("2 - NO, write to .txt")
	reader := bufio.NewReader(os.Stdin)
	r, _, err := reader.ReadRune()
	if err != nil {
		panic(err)
	}
	if r == bytes.Runes([]byte("1"))[0] {
		return true
	}
	return false
}

//convert any file with .xml ext in dir
func processDir(dir string) {
	files, err := ioutil.ReadDir(dir)
	if err != nil {
		log.Fatal(err)
		return
	}
	for _, f := range files {
		filename := f.Name()
		dim := filepath.Ext(filename)
		//fmt.Println(dim)
		if dim == ".xml" {
			convert(f.Name(), askForGetParams(), askForXlsx())
		}
	}
}

//convert file
func convert(filename string, getParams bool, toxlsx bool) {
	fmt.Println("statring " + filename)
	catalog, err := getCatalog(filename)
	if err != nil {
		fmt.Println(err.Error())
		return
	}
	if toxlsx {
		xlsx := excelize.NewFile()
		writeTableToXlsx(xlsx, "offers", catalog.GetOfferTable())
		writeTableToXlsx(xlsx, "categories", catalog.GetCategoryTable())
		writeTableToXlsx(xlsx, "categoryTree", catalog.GetCategoryTreeTable())
		if getParams {
			writeTableToXlsx(xlsx, "params", catalog.GetParamsTable())
		}
		xlsx.DeleteSheet("Sheet1")
		newfilename := filename[0:len(filename)-len(filepath.Ext(filename))] + ".xlsx"
		fmt.Println("saving " + newfilename)
		xlsx.SaveAs(newfilename)
		fmt.Println(newfilename + " created")
	} else {
		writeTableToCsv(filename[0:len(filename)-len(filepath.Ext(filename))]+"_offers.txt", catalog.GetOfferTable())
		writeTableToCsv(filename[0:len(filename)-len(filepath.Ext(filename))]+"_categories.txt", catalog.GetCategoryTable())
		writeTableToCsv(filename[0:len(filename)-len(filepath.Ext(filename))]+"_categoryTree.txt", catalog.GetCategoryTreeTable())
		if getParams {
			writeTableToCsv(filename[0:len(filename)-len(filepath.Ext(filename))]+"_params.csv", catalog.GetParamsTable())
		}
	}

}

//parse YmlCatalog object from file
func getCatalog(filename string) (Yml_catalog, error) {
	catalog, err := decodeCatalog(filename, charmap.Windows1251.NewDecoder())
	if err != nil {
		catalog, err = decodeCatalog(filename, nil)
	}
	return catalog, err
}

//parse file using given decoder or without char mapping if nil
func decodeCatalog(filename string, decoder *encoding.Decoder) (Yml_catalog, error) {
	doc := Yml_catalog{}
	xmlFile, err := os.Open(filename)
	if err != nil {
		return doc, err
	}
	defer xmlFile.Close()
	b := xml.NewDecoder(xmlFile)
	if decoder != nil {
		b.CharsetReader = func(charset string, input io.Reader) (io.Reader, error) {
			switch charset {
			case "Windows-1251":
				fallthrough
			case "windows-1251":
				return decoder.Reader(input), nil
			default:
				return nil, fmt.Errorf("unknown charset: %s", charset)
			}
		}
	} else {
		b.CharsetReader = func(charset string, input io.Reader) (io.Reader, error) {
			switch charset {
			case "Windows-1251":
				fallthrough
			case "windows-1251":
				return input, nil
			default:
				return nil, fmt.Errorf("unknown charset: %s", charset)
			}
		}
	}

	err = b.Decode(&doc)
	if err != nil {
		return doc, err
	}
	return doc, nil
}

//write Table object content to given xlsx file into sheet with given name
func writeTableToXlsx(xlsx *excelize.File, sheetname string, table Table) {
	fmt.Println("\t" + "writing sheet \"" + sheetname + "\"...")
	xlsx.NewSheet(sheetname)
	for k, v := range table.Columns {
		columnname := getColumnName(v)
		xlsx.SetCellValue(sheetname, columnname+"1", k)
	}
	var i int

	bar := pb.StartNew(len(table.Rows))
	for k, v := range table.Rows {
		i++
		bar.Increment()
		rowname := strconv.Itoa(k + 2)

		//xlsx.SetSheetRow(sheetname, "A" + rowname,&sl)
		for kk, vv := range v.Cells {
			columnname := getColumnName(kk)
			xlsx.SetCellValue(sheetname, columnname+rowname, vv)
		}

	}
	bar.Finish()
}

//write Table object content to given xlsx file into sheet with given name
func writeTableToCsv(filename string, table Table) {
	fmt.Println("\t" + "writing file \"" + filename + "\"...")

	file, err := os.OpenFile(filename, os.O_APPEND|os.O_WRONLY, os.ModeAppend)
	if err != nil {
		file, err = os.Create(filename)
	}
	file.Close()

	fileHandle, _ := os.OpenFile(filename, os.O_APPEND, 0666)
	writer := bufio.NewWriter(fileHandle)
	defer fileHandle.Close()

	maxcol := 0
	colMap := make(map[int]string)
	for colName, colNum := range table.Columns {
		colMap[colNum] = colName
		if maxcol < colNum {
			maxcol = colNum
		}
	}
	data := ""
	for i := 0; i < maxcol; i++ {
		data += colMap[i] + "\t"
	}
	data += "\r\n"
	fmt.Fprint(writer, data)

	bar := pb.StartNew(len(table.Rows))
	for _, row := range table.Rows {
		maxcol := 0
		for _, colNum := range table.Columns {
			if _, ok := row.Cells[colNum]; ok {
				if maxcol < colNum {
					maxcol = colNum
				}
			}
		}
		data := ""
		for i := 0; i < maxcol; i++ {
			data += row.Cells[i] + "\t"
		}
		data += "\r\n"
		fmt.Fprint(writer, data)
		//file.Sync()
		bar.Increment()
	}
	bar.Finish()
}

//function to get column name by column number. supports up to 17526 columns
func getColumnName(v int) string {
	var columnname string
	if v <= 17526 {
		alfabet := []string{"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"}
		columnnum := v
		if columnnum >= len(alfabet) {
			if columnnum >= len(alfabet)*len(alfabet) {
				first := (columnnum - (columnnum % (len(alfabet) * len(alfabet)))) / (len(alfabet) * len(alfabet))
				last := (columnnum - first) % len(alfabet)
				mid := (columnnum - (first * len(alfabet) * len(alfabet)) - last) / len(alfabet)
				columnname = alfabet[first] + alfabet[mid] + alfabet[last]
			} else {
				last := alfabet[columnnum%len(alfabet)]
				first := alfabet[columnnum/len(alfabet)-1]
				columnname = first + last
			}
		} else {
			first := alfabet[columnnum]
			columnname = first
		}
	} else {
		columnname = getColumnName(17526)
	}
	return columnname
}
