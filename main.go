package main

import (
	"encoding/csv"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
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

func div(Num int) string {
	var (
		Str  string = ""
		k    int
		temp []int //保存转化后每一位数据的值，然后通过索引的方式匹配A-Z
	)
	//用来匹配的字符A-Z
	Slice := []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
		"P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

	if Num > 26 { //数据大于26需要进行拆分
		for {
			k = Num % 26 //从个位开始拆分，如果求余为0，说明末尾为26，也就是Z，如果是转化为26进制数，则末尾是可以为0的，这里必须为A-Z中的一个
			if k == 0 {
				temp = append(temp, 26)
				k = 26
			} else {
				temp = append(temp, k)
			}
			Num = (Num - k) / 26 //减去Num最后一位数的值，因为已经记录在temp中
			if Num <= 26 {       //小于等于26直接进行匹配，不需要进行数据拆分
				temp = append(temp, Num)
				break
			}
		}
	} else {
		return Slice[Num]
	}

	for _, value := range temp {
		Str = Slice[value] + Str //因为数据切分后存储顺序是反的，所以Str要放在后面
	}
	return Str
}
func toXlsx(data [][]string, outFile string) {
	var axis []string
	for i := 0; i < len(data[0]); i++ {
		axis = append(axis, div(i+1))
	}

	//log.Println("长度:", len(data))

	f := excelize.NewFile()
	// Create a new sheet.
	index := f.NewSheet("Sheet1")

	for line := range data {
		for key := range data[line] {
			f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", axis[key], line), data[line][key])
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
