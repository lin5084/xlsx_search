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
		fmt.Printf("✅ 一共找到 %d 个待检索文档\n", len(otherFiles))
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

		fmt.Printf("\n👉 正在检索关键词【%s】，行号：%d\n", key, row)

		var foundLines [][]string
		for _, of := range otherFiles {
			baseName := filepath.Base(of)
			fmt.Printf("   正在读取文档：%s\n", baseName)

			fOther, err := excelize.OpenFile(of)
			if err != nil {
				fmt.Printf("   ⚠️ 跳过文件 %s，打开失败：%v\n", of, err)
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
						colName, err := excelize.ColumnNumberToName(oc)
						if err != nil {
							break
						}
						v, _ := fOther.GetCellValue(osheet, colName+fmt.Sprintf("%d", or))
						line = append(line, v)
						if oc >= 2 && line[oc-1] == "" && line[oc-2] == "" {
							break
						}
						oc++
					}
					// ✅ 在本行末尾追加来源文档名称
					line = append(line, "来源:"+baseName)
					foundLines = append(foundLines, line)
					fmt.Printf("      ✅ 在 %s 找到匹配行！\n", baseName)
				}
				or++
			}
			fOther.Close()
		}

		if len(foundLines) > 0 {
			col := 2
			for _, line := range foundLines {
				for ci, val := range line {
					colName, err := excelize.ColumnNumberToName(col + ci)
					if err != nil {
						continue
					}
					cellName := colName + fmt.Sprintf("%d", row)
					_ = fTarget.SetCellValue(sheetName, cellName, val)
				}
				col += len(line)
			}
		} else {
			_ = fTarget.SetCellValue(sheetName, "B"+fmt.Sprintf("%d", row), "不存在")
			fmt.Printf("   ❌ 全部文档都没有找到【%s】\n", key)
		}
		row++
	}

	err = fTarget.Save()
	if err != nil {
		fmt.Println("保存待查找.xlsx失败:", err)
		waitEnter()
		return
	}

	fmt.Println("\n🎉 全部处理完成！已保存到【待查找.xlsx】")
	waitEnter()
}

func waitEnter() {
	fmt.Println("\n按回车键退出...")
	_, _ = fmt.Scanln()
}
