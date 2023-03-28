package main

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx/v3"
)

func main() {
	// 检查命令行参数
	if len(os.Args) != 3 {
		fmt.Println("Usage: xlsx2lua <input.xlsx> <output.lua>")
		return
	}
	inputFilename := os.Args[1]
	outputFilename := os.Args[2]

	// 打开 Excel 文件
	f, err := xlsx.OpenFile(inputFilename)
	if err != nil {
		fmt.Println(err)
		return
	}

	// 获取第一个 Sheet
	sh := f.Sheets[0]

	// 获取列数和行数
	numCols := sh.MaxCol
	numRows := sh.MaxRow

	// 获取列名
	colNames := make([]string, numCols)
	for colIndex := 0; colIndex < numCols; colIndex++ {
		cell, _ := sh.Cell(0, colIndex)
		if cell.Value == "" {
			continue
		}
		colNames[colIndex] = cell.Value
	}

	// 构造表格数据
	var data string
	data += "{\n"
	for rowIndex := 1; rowIndex < numRows; rowIndex++ {
		// 检查该行是否为空行
		isEmptyRow := true
		for colIndex := 0; colIndex < numCols; colIndex++ {
			cell, _ := sh.Cell(rowIndex, colIndex)
			if cell.Value != "" {
				isEmptyRow = false
				break
			}
		}
		if isEmptyRow {
			continue
		}

		data += "    {\n"
		for colIndex := 0; colIndex < numCols; colIndex++ {
			if colNames[colIndex] == "" {
				continue
			}

			cell, _ := sh.Cell(rowIndex, colIndex)
			value := cell.Value

			// 根据单元格的类型进行转换
			switch cell.Type() {
			case xlsx.CellTypeNumeric:
				if strings.Contains(value, ".") {
					f, _ := strconv.ParseFloat(value, 64)
					value = fmt.Sprintf("%f", f)
				} else {
					i, _ := strconv.ParseInt(value, 10, 64)
					value = fmt.Sprintf("%d", i)
				}
			case xlsx.CellTypeBool:
				value = fmt.Sprintf("%t", cell.Bool())
			}

			data += fmt.Sprintf("        %s = %q,\n", colNames[colIndex], value)
		}
		data += "    },\n"
	}
	data += "}"

	// 将数据输出到文件
	outputFile, err := os.Create(outputFilename)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer outputFile.Close()

	fmt.Fprint(outputFile, "return ")
	fmt.Fprint(outputFile, data)

	fmt.Println("转换成功！")
}
