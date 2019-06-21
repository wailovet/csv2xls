package main

import (
	"encoding/csv"
	"fmt"
	"github.com/tealeg/xlsx"
	"log"
	"os"
	"path/filepath"
)

func main() {
	if len(os.Args) < 3 {
		fmt.Println("参数错误: 运行示例:csv2xls <csv文件路径> <要保存的xlsx文件路径>")
		os.Exit(0)
	}

	inFile := os.Args[1]
	inFile, _ = filepath.Abs(inFile)

	outFile := os.Args[2]
	outFile, _ = filepath.Abs(outFile)

	to(inFile, outFile)
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
