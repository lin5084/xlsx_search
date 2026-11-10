package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Record struct {
	Data      []string
	Src       string
	SourceRow int
}

type KeyItem struct {
	Key     string
	Records []Record
}

// dedupRecords 外部记录去重：来源文件名+行号
func dedupRecords(recs []Record) []Record {
	seen := make(map[string]bool)
	var out []Record
	for _, r := range recs {
		k := fmt.Sprintf("%s|%d", r.Src, r.SourceRow)
		if seen[k] {
			fmt.Printf("      ⚠️  去重跳过重复外部记录：%s,行%d\n", r.Src, r.SourceRow)
			continue
		}
		seen[k] = true
		out = append(out, r)
	}
	return out
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

	// 打开待查找文档读取A1作为表头匹配标记
	fTargetTmp, err := excelize.OpenFile(targetPath)
	if err != nil {
		fmt.Println("打开待查找.xlsx失败(A1读取)", err)
		waitEnter()
		return
	}
	sheetsTmp := fTargetTmp.GetSheetList()
	if len(sheetsTmp) == 0 {
		fmt.Println("待查找.xlsx无工作表")
		fTargetTmp.Close()
		waitEnter()
		return
	}
	sheetTmp := sheetsTmp[0]
	headerMarker, err := fTargetTmp.GetCellValue(sheetTmp, "A1")
	fTargetTmp.Close()
	if err != nil || strings.TrimSpace(headerMarker) == "" {
		fmt.Println("❌ 待查找.xlsx A1单元格不能为空，请设置表头标记！")
		waitEnter()
		return
	}
	headerMarker = strings.TrimSpace(headerMarker)
	fmt.Printf("✅ 使用待查找.xlsx A1作为表头标记：【%s】\n", headerMarker)

	lookupMap := make(map[string][]Record)
	var missingHeaderDocs []string // 收集：表头找不到标记的文档

	for _, of := range otherFiles {
		baseName := filepath.Base(of)
		fmt.Printf("📖 正在加载文档：%s\n", baseName)

		fOther, err := excelize.OpenFile(of)
		if err != nil {
			fmt.Printf("⚠️ 跳过文件 %s，打开失败：%v\n", baseName, err)
			missingHeaderDocs = append(missingHeaderDocs, baseName)
			continue
		}
		osheets := fOther.GetSheetList()
		if len(osheets) == 0 {
			fOther.Close()
			missingHeaderDocs = append(missingHeaderDocs, baseName)
			continue
		}
		osheet := osheets[0]

		// -------- 在第1行表头寻找匹配标记的列索引 --------
		targetSearchCol := -1
		colScan := 1
		for {
			colHName, err := excelize.ColumnNumberToName(colScan)
			if err != nil {
				break
			}
			cellH, err := fOther.GetCellValue(osheet, colHName+"1")
			if err != nil {
				break
			}
			cellH = strings.TrimSpace(cellH)
			if cellH == headerMarker {
				targetSearchCol = colScan
				break
			}
			// 表头连续两个空，停止扫描表头行
			if colScan >= 2 {
				prevColName, err := excelize.ColumnNumberToName(colScan - 1)
				if err != nil {
					break
				}
				prevVal, _ := fOther.GetCellValue(osheet, prevColName+"1")
				if strings.TrimSpace(cellH) == "" && strings.TrimSpace(prevVal) == "" {
					break
				}
			}
			colScan++
		}

		if targetSearchCol == -1 {
			fmt.Printf("   ⚠️ 文档 %s 第一行表头未找到标记【%s】，加入清单\n", baseName, headerMarker)
			missingHeaderDocs = append(missingHeaderDocs, baseName)
			fOther.Close()
			continue
		}
		fmt.Printf("   ✅ 在第%d列找到表头标记，将以此列检索\n", targetSearchCol)

		// 遍历数据行（从第2行开始，跳过表头）
		or := 2
		for {
			searchColName, err := excelize.ColumnNumberToName(targetSearchCol)
			if err != nil {
				break
			}
			searchCellVal, err := fOther.GetCellValue(osheet, searchColName+fmt.Sprintf("%d", or))
			if err != nil {
				break
			}
			searchCellVal = strings.TrimSpace(searchCellVal)
			if searchCellVal == "" {
				break
			}

			// 按空格分割单元格内容，得到片段数组
			tokens := strings.Fields(searchCellVal)

			// 读取整行全部单元格
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
			line = append(line, "来源:"+baseName)
			line = append(line, fmt.Sprintf("来源行号:%d", or))

			// 将每一个空格分割片段作为检索key存入map
			for _, tk := range tokens {
				tkTrim := strings.TrimSpace(tk)
				if tkTrim == "" {
					continue
				}
				lookupMap[tkTrim] = append(lookupMap[tkTrim], Record{
					Data:      line,
					Src:       baseName,
					SourceRow: or,
				})
			}
			or++
		}
		fOther.Close()
	}
	fmt.Printf("✅ 所有文档加载完成，一共缓存 %d 个检索关键词\n", len(lookupMap))
	fmt.Printf("⚠️ 表头缺失标记的文档数量：%d\n", len(missingHeaderDocs))

	// ========= 读取待查找A列待检索关键词（跳过A1表头，从第2行读取） =========
	fmt.Println("\n📋 读取待查找列表（跳过A1表头）...")
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
	readRow := 2 // A1是表头标记，从第2行读取待检索关键词
	seenTargetKey := make(map[string]bool)
	skipCount := 0

	for {
		key, err := fTarget.GetCellValue(sheetName, fmt.Sprintf("A%d", readRow))
		if err != nil {
			break
		}
		key = strings.TrimSpace(key)
		if key == "" {
			break
		}

		if seenTargetKey[key] {
			fmt.Printf("⚠️ 待查找表A列检测到重复关键词【%s】，跳过该行\n", key)
			skipCount++
			readRow++
			continue
		}
		seenTargetKey[key] = true

		recs := lookupMap[key]
		recs = dedupRecords(recs)
		keyList = append(keyList, KeyItem{
			Key:     key,
			Records: recs,
		})
		readRow++
	}
	fmt.Printf("✅ 原始读取：%d行，跳过待查找表重复关键词：%d行，实际处理关键词：%d个\n", readRow-2, skipCount, len(keyList))

	// 【重要修改】不再清空A1！输出结果从第2行开始，maxRow初始值=2
	maxRow := 2
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

	// ========= 输出：表头没有A1标记的文档清单，接续在后面 =========
	fmt.Println("\n📋 输出表头缺失标记的文档清单...")
	_ = fTarget.SetCellValue(sheetName, "A"+fmt.Sprintf("%d", maxRow), "下面是文档表头没有A1单元格数据的文档")
	maxRow++
	for _, docName := range missingHeaderDocs {
		_ = fTarget.SetCellValue(sheetName, "A"+fmt.Sprintf("%d", maxRow), docName)
		maxRow++
	}

	err = fTarget.Save()
	if err != nil {
		fmt.Println("保存待查找.xlsx失败:", err)
		waitEnter()
		return
	}

	fmt.Printf("\n🎉 处理完毕！\n✅ 匹配总行数：%d\n❌ 无匹配条目：%d\n缺失表头文档数量：%d\n",
		matchCount, noMatchCount, len(missingHeaderDocs))
	waitEnter()
}

func waitEnter() {
	fmt.Println("\n按回车键退出...")
	_, _ = fmt.Scanln()
}
