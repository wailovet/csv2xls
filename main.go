package main

import (
	"encoding/csv"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/tealeg/xlsx"
	"github.com/wailovet/osmanthuswine"
	"github.com/wailovet/osmanthuswine/src/core"
	"log"
	"os"
	"path/filepath"
)

type Csv2xlsx struct {
	core.Controller
}

func (that *Csv2xlsx) ToXlsx() {

	inFile := that.Request.REQUEST["inFile"]
	inFile, _ = filepath.Abs(inFile)

	outFile := that.Request.REQUEST["outFile"]
	outFile, _ = filepath.Abs(outFile)
	to(inFile, outFile)
}

func main() {

	core.GetInstanceRouterManage().Registered(&Csv2xlsx{})
	core.GetInstanceConfig().Port = "17003"
	osmanthuswine.Run()

	if len(os.Args) < 3 {
		fmt.Println("参数错误: 运行示例:csv2xls <csv文件路径> <要保存的xlsx文件路径>")
		os.Exit(0)
	}

	inFile := os.Args[1]
	inFile, _ = filepath.Abs(inFile)

	outFile := os.Args[2]
	outFile, _ = filepath.Abs(outFile)

}

func to(inFile string, outFile string) {
	csvFile, err := os.Open(inFile)
	if err != nil {
		log.Fatal(err)
	}
	defer csvFile.Close()

	csvReader := csv.NewReader(csvFile)
	data, err := csvReader.ReadAll()
	if err != nil {
		log.Fatal(err)
	}

	toXlsx(data, outFile)
}

func toXlsx(data [][]string, outFile string) {

	//log.Println("长度:", len(data))

	f := excelize.NewFile()
	// Create a new sheet.
	index := f.NewSheet("Sheet1")

	for line := range data {
		for key := range data[line] {
			keyc := rune(int('A') + key)
			f.SetCellValue("Sheet1", fmt.Sprintf("%c%d", keyc, line), data[line][key])
		}
	}

	// Set active sheet of the workbook.
	f.SetActiveSheet(index)
	// Save xlsx file by the given path.

	err := f.SaveAs(outFile)
	if err != nil {
		log.Println(err)
	}
}
func toXlsx2(data [][]string, outFile string) {
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}

	for line := range data {
		row = sheet.AddRow()
		for key := range data[line] {
			cell = row.AddCell()
			cell.Value = data[line][key]
		}
	}

	err = file.Save(outFile)
	if err != nil {
		fmt.Printf(err.Error())
	}
}
