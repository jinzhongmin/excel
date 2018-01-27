package excel

import (
	"log"
	"strconv"
)

//Const ..
type Const struct {
	FILEFORMATxlAddIn                       int
	FILEFORMATxlAddIn8                      int
	FILEFORMATxlCSV                         int
	FILEFORMATxlCSVMac                      int
	FILEFORMATxlCSVMSDOS                    int
	FILEFORMATxlCSVWindows                  int
	FILEFORMATxlCurrentPlatformText         int
	FILEFORMATxlDBF2                        int
	FILEFORMATxlDBF3                        int
	FILEFORMATxlDBF4                        int
	FILEFORMATxlDIF                         int
	FILEFORMATxlExcel12                     int
	FILEFORMATxlExcel2                      int
	FILEFORMATxlExcel2FarEast               int
	FILEFORMATxlExcel3                      int
	FILEFORMATxlExcel4                      int
	FILEFORMATxlExcel4Workbook              int
	FILEFORMATxlExcel5                      int
	FILEFORMATxlExcel7                      int
	FILEFORMATxlExcel8                      int
	FILEFORMATxlExcel9795                   int
	FILEFORMATxlHTML                        int
	FILEFORMATxlIntlAddIn                   int
	FILEFORMATxlIntlMacro                   int
	FILEFORMATxlOpenDocumentSpreadsheet     int
	FILEFORMATxlOpenXMLAddIn                int
	FILEFORMATxlOpenXMLStrictWorkbook       int
	FILEFORMATxlOpenXMLTemplate             int
	FILEFORMATxlOpenXMLTemplateMacroEnabled int
	FILEFORMATxlOpenXMLWorkbook             int
	FILEFORMATxlOpenXMLWorkbookMacroEnabled int
	FILEFORMATxlSYLK                        int
	FILEFORMATxlTemplate                    int
	FILEFORMATxlTemplate8                   int
	FILEFORMATxlTextMac                     int
	FILEFORMATxlTextMSDOS                   int
	FILEFORMATxlTextPrinter                 int
	FILEFORMATxlTextWindows                 int
	FILEFORMATxlUnicodeText                 int
	FILEFORMATxlWebArchive                  int
	FILEFORMATxlWJ2WD1                      int
	FILEFORMATxlWJ3                         int
	FILEFORMATxlWJ3FJ3                      int
	FILEFORMATxlWK1                         int
	FILEFORMATxlWK1ALL                      int
	FILEFORMATxlWK1FMT                      int
	FILEFORMATxlWK3                         int
	FILEFORMATxlWK3FM3                      int
	FILEFORMATxlWK4                         int
	FILEFORMATxlWKS                         int
	FILEFORMATxlWorkbookDefault             int
	FILEFORMATxlWorkbookNormal              int
	FILEFORMATxlWorks2FarEast               int
	FILEFORMATxlWQ1                         int
	FILEFORMATxlXMLSpreadsheet              int
}

//NewConst ..
func NewConst() *Const {
	c := new(Const)

	c.FILEFORMATxlAddIn = 18
	c.FILEFORMATxlAddIn8 = 18
	c.FILEFORMATxlCSV = 6
	c.FILEFORMATxlCSVMac = 22
	c.FILEFORMATxlCSVMSDOS = 24
	c.FILEFORMATxlCSVWindows = 23
	c.FILEFORMATxlCurrentPlatformText = -4158
	c.FILEFORMATxlDBF2 = 7
	c.FILEFORMATxlDBF3 = 8
	c.FILEFORMATxlDBF4 = 11
	c.FILEFORMATxlDIF = 9
	c.FILEFORMATxlExcel12 = 50
	c.FILEFORMATxlExcel2 = 16
	c.FILEFORMATxlExcel2FarEast = 27
	c.FILEFORMATxlExcel3 = 29
	c.FILEFORMATxlExcel4 = 33
	c.FILEFORMATxlExcel4Workbook = 35
	c.FILEFORMATxlExcel5 = 39
	c.FILEFORMATxlExcel7 = 39
	c.FILEFORMATxlExcel8 = 56
	c.FILEFORMATxlExcel9795 = 43
	c.FILEFORMATxlHTML = 44
	c.FILEFORMATxlIntlAddIn = 26
	c.FILEFORMATxlIntlMacro = 25
	c.FILEFORMATxlOpenDocumentSpreadsheet = 60
	c.FILEFORMATxlOpenXMLAddIn = 55
	c.FILEFORMATxlOpenXMLStrictWorkbook = 61
	c.FILEFORMATxlOpenXMLTemplate = 54
	c.FILEFORMATxlOpenXMLTemplateMacroEnabled = 53
	c.FILEFORMATxlOpenXMLWorkbook = 51
	c.FILEFORMATxlOpenXMLWorkbookMacroEnabled = 52
	c.FILEFORMATxlSYLK = 2
	c.FILEFORMATxlTemplate = 17
	c.FILEFORMATxlTemplate8 = 17
	c.FILEFORMATxlTextMac = 19
	c.FILEFORMATxlTextMSDOS = 21
	c.FILEFORMATxlTextPrinter = 36
	c.FILEFORMATxlTextWindows = 20
	c.FILEFORMATxlUnicodeText = 42
	c.FILEFORMATxlWebArchive = 45
	c.FILEFORMATxlWJ2WD1 = 14
	c.FILEFORMATxlWJ3 = 40
	c.FILEFORMATxlWJ3FJ3 = 41
	c.FILEFORMATxlWK1 = 5
	c.FILEFORMATxlWK1ALL = 31
	c.FILEFORMATxlWK1FMT = 30
	c.FILEFORMATxlWK3 = 15
	c.FILEFORMATxlWK3FM3 = 32
	c.FILEFORMATxlWK4 = 38
	c.FILEFORMATxlWKS = 4
	c.FILEFORMATxlWorkbookDefault = 51
	c.FILEFORMATxlWorkbookNormal = -4143
	c.FILEFORMATxlWorks2FarEast = 28
	c.FILEFORMATxlWQ1 = 34
	c.FILEFORMATxlXMLSpreadsheet = 46

	return c
}

//NumToLetter ..
func NumToLetter(col, row int) string {
	if col > 256 {
		log.Fatalln("in func NumToLetter : col must <= 256")
	}
	c := string(rune(col%26 + 64))
	if col/26 > 0 {
		c = string(rune(col/26+64)) + c
	}
	r := strconv.Itoa(row)

	return c + r
}

//LetterToNum ..
func LetterToNum(cr string) (int, int) {
	var err error
	point := []byte(cr)
	col := int(point[0] - 64)
	row := 0
	if int(point[1]) > 64 && int(point[0]) < 91 {
		col = col*26 + int(point[1]-64)
		row, err = strconv.Atoi(string(point[2:]))
		if err != nil {
			log.Fatalln("in func LetterToNum : arg err")
		}
		if col > 256 {
			log.Fatalln("in func LetterToNum : col must <= 256")
		}
		return col, row
	}
	row, err = strconv.Atoi(string(point[1:]))
	if err != nil {
		log.Fatalln("in func LetterToNum : arg err")
	}
	if col > 256 {
		log.Fatalln("in func LetterToNum : col must <= 256")
	}
	return col, row
}
