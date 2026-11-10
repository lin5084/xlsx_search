package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Record 增加 SourceRow 来源文档里的行号
type Record struct {
	Data      []string
	Src       string
	SourceRow int
}

type KeyItem struct {
	Key     string
	Records []Record
}

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
		fmt.Printf("✅ 一共找到 %d 个待检索文档，开始预加载...\n", len(otherFiles))
	}

	lookupMap := make(map[string][]Record)
	for _, of := range otherFiles {
		baseName := filepath.Base(of)
		fmt.Printf("📖 正在加载文档：%s\n", baseName)

		fOther, err := excelize.OpenFile(of)
		if err != nil {
			fmt.Printf("⚠️ 跳过文件 %s，打开失败：%v\n", baseName, err)
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
			cellVal, err := fOther.GetCellValue(osheet, fmt.Sprintf("A%d", or))
			if err != nil {
				break
			}
			cellVal = strings.TrimSpace(cellVal)
			if cellVal == "" {
				break
			}

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
			// 先追加来源文件名，再追加来源行号
			line = append(line, "来源:"+baseName)
			line = append(line, fmt.Sprintf("来源行号:%d", or))

			lookupMap[cellVal] = append(lookupMap[cellVal], Record{
				Data:      line,
				Src:       baseName,
				SourceRow: or,
			})
			or++
		}
		fOther.Close()
	}
	fmt.Printf("✅ 所有文档加载完成，一共缓存 %d 个唯一关键词\n", len(lookupMap))

	fmt.Println("\n📋 读取待查找列表...")
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

	var keyList []KeyItem
	readRow := 1
	for {
		key, err := fTarget.GetCellValue(sheetName, fmt.Sprintf("A%d", readRow))
		if err != nil {
			break
		}
		key = strings.TrimSpace(key)
		if key == "" {
			break
		}
		recs := lookupMap[key]
		keyList = append(keyList, KeyItem{
			Key:     key,
			Records: recs,
		})
		readRow++
	}
	fmt.Printf("✅ 读取到 %d 个待检索关键词\n", len(keyList))

	_ = fTarget.SetSheetRow(sheetName, "A1", nil)
	maxRow := 1
	matchCount := 0
	noMatchCount := 0

	fmt.Println("\n===== 开始生成结果 =====")
	for _, item := range keyList {
		key := item.Key
		recs := item.Records
		if len(recs) == 0 {
			fmt.Printf("🔍 关键词【%s】 → ❌ 不存在\n", key)
			_ = fTarget.SetCellValue(sheetName, "A"+fmt.Sprintf("%d", maxRow), key)
			_ = fTarget.SetCellValue(sheetName, "B"+fmt.Sprintf("%d", maxRow), "不存在")
			maxRow++
			noMatchCount++
			continue
		}
		fmt.Printf("🔍 关键词【%s】 → ✅ 找到 %d 条\n", key, len(recs))
		for _, rec := range recs {
			fmt.Printf("      ▶ 来源：%s，来源行号：%d\n", rec.Src, rec.SourceRow)
			_ = fTarget.SetCellValue(sheetName, "A"+fmt.Sprintf("%d", maxRow), key)
			for ci, val := range rec.Data {
				colIdx := 2 + ci
				colName, err := excelize.ColumnNumberToName(colIdx)
				if err != nil {
					continue
				}
				cell := colName + fmt.Sprintf("%d", maxRow)
				_ = fTarget.SetCellValue(sheetName, cell, val)
			}
			maxRow++
			matchCount++
		}
	}

	err = fTarget.Save()
	if err != nil {
		fmt.Println("保存待查找.xlsx失败:", err)
		waitEnter()
		return
	}

	fmt.Printf("\n🎉 处理完毕！\n✅ 匹配总行数：%d\n❌ 无匹配条目：%d\n", matchCount, noMatchCount)
	waitEnter()
}

func waitEnter() {
	fmt.Println("\n按回车键退出...")
	_, _ = fmt.Scanln()
}
