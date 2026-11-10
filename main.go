package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	cwd, err := os.Getwd()
	if err != nil {
		fmt.Println("获取当前目录失败:", err)
		waitEnter()
		return
	}

	targetName := "待查找.xlsx"
	targetPath := filepath.Join(cwd, targetName)
	if _, err := os.Stat(targetPath); os.IsNotExist(err) {
		fmt.Printf("错误：找不到 %s，请放在同目录！\n", targetName)
		waitEnter()
		return
	}

	// 收集目录下其他xlsx
	var otherFiles []string
	entries, err := filepath.Glob(filepath.Join(cwd, "*.xlsx"))
	if err != nil {
		fmt.Println("读取目录失败:", err)
		waitEnter()
		return
	}
	for _, fp := range entries {
		if filepath.Base(fp) != targetName {
			otherFiles = append(otherFiles, fp)
		}
	}
	if len(otherFiles) == 0 {
		fmt.Println("警告：目录里没有其他xlsx文件用来检索")
	} else {
		fmt.Printf("找到待检索文件数量：%d\n", len(otherFiles))
	}

	fTarget, err := excelize.OpenFile(targetPath)
	if err != nil {
		fmt.Println("打开待查找.xlsx失败:", err)
		waitEnter()
		return
	}
	defer fTarget.Close()

	sheets := fTarget.GetSheetList()
	if len(sheets) == 0 {
		fmt.Println("待查找.xlsx没有工作表！")
		waitEnter()
		return
	}
	sheetName := sheets[0]

	row := 1
	for {
		key, err := fTarget.GetCellValue(sheetName, fmt.Sprintf("A%d", row))
		if err != nil {
			break
		}
		key = strings.TrimSpace(key)
		if key == "" {
			break
		}

		var foundLines [][]string
		for _, of := range otherFiles {
			fOther, err := excelize.OpenFile(of)
			if err != nil {
				fmt.Printf("跳过文件 %s，打开失败：%v\n", of, err)
				continue
			}
			osheets := fOther.GetSheetList()
			if len(osheets) == 0 {
				fOther.Close()
				continue
			}
			osheet := osheets[0]

			or := 1
			for {
				cellVal, _ := fOther.GetCellValue(osheet, fmt.Sprintf("A%d", or))
				cellVal = strings.TrimSpace(cellVal)
				if cellVal == "" {
					break
				}
				if cellVal == key {
					var line []string
					oc := 1
					for {
						colName := excelize.ColumnNumberToName(oc)
						v, _ := fOther.GetCellValue(osheet, colName+fmt.Sprintf("%d", or))
						line = append(line, v)
						// 连续两个空就停止读这一行
						if oc >= 2 && line[oc-1] == "" && line[oc-2] == "" {
							break
						}
						oc++
					}
					foundLines = append(foundLines, line)
				}
				or++
			}
			fOther.Close()
		}

		if len(foundLines) > 0 {
			col := 2
			for _, line := range foundLines {
				for ci, val := range line {
					cellName := excelize.ColumnNumberToName(col+ci) + fmt.Sprintf("%d", row)
					_ = fTarget.SetCellValue(sheetName, cellName, val)
				}
				col += len(line)
			}
		} else {
			_ = fTarget.SetCellValue(sheetName, "B"+fmt.Sprintf("%d", row), "不存在")
		}
		row++
	}

	err = fTarget.Save()
	if err != nil {
		fmt.Println("保存待查找.xlsx失败:", err)
		waitEnter()
		return
	}

	fmt.Println("✅ 全部处理完成！")
	waitEnter()
}

func waitEnter() {
	fmt.Println("\n按回车键退出...")
	_, _ = fmt.Scanln()
}
