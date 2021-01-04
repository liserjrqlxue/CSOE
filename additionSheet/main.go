package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/liserjrqlxue/goUtil/textUtil"
)

// flag
var (
	workDir = flag.String(
		"workDir",
		"",
		"work dir",
	)
	final = flag.String(
		"final",
		"",
		"final.result.xlsx",
	)
	ma = flag.String(
		"ma",
		"",
		"MA result, default is -workDir/Y159_MA_update/MA_update.xls",
	)
	cnv = flag.String(
		"cnv",
		"",
		"CAH_GBA CNV result, default is -workDir/CAH_GBA/CNV/CNV_report.reform.xlsx",
	)
	gba = flag.String(
		"gba",
		"",
		"gba.PE_SE.xlsx, default is -workDir/CAH_GBA/$(basename -workDir).gba.PE_SE.xlsx",
	)
	com = flag.String(
		"com",
		"",
		"com_snp.xlsx, default is -workDir/CAH_GBA/$(basename -workDir).com_snp.xlsx",
	)
)

func main() {
	flag.Parse()
	if *workDir != "" {
		*workDir = simpleUtil.HandleError(filepath.Abs(*workDir)).(string)
		var baseDir = filepath.Base(*workDir)
		if *ma == "" {
			*ma = filepath.Join(*workDir, "Y159_MA_update", "MA_update.xls")
		}
		if *cnv == "" {
			*cnv = filepath.Join(*workDir, "CAH_GBA", "CNV", "CNV_report.reform.xlsx")
		}
		if *gba == "" {
			*gba = filepath.Join(*workDir, "CAH_GBA", baseDir+".gba.PE_SE.xlsx")
		}
		if *com == "" {
			*com = filepath.Join(*workDir, "CAH_GBA", baseDir+".com_snp.xlsx")
		}
	}
	if *ma == "" || *cnv == "" || *gba == "" || *com == "" {
		flag.Usage()
		fmt.Println("-workDir or -ma/-cnv/-gba/-com are required!")
		os.Exit(1)
	}
	if *final == "" {
		flag.Usage()
		fmt.Println("-final is required!")
		os.Exit(1)
	}
	var maSlice = textUtil.File2Slice(*ma, "\t")
	var finalXlsx = simpleUtil.HandleError(excelize.OpenFile(*final)).(*excelize.File)
	var cnvXlsx = simpleUtil.HandleError(excelize.OpenFile(*cnv)).(*excelize.File)
	var gbaXlsx = simpleUtil.HandleError(excelize.OpenFile(*gba)).(*excelize.File)
	var comXlsx = simpleUtil.HandleError(excelize.OpenFile(*com)).(*excelize.File)
	AppendSheet(cnvXlsx, finalXlsx, "CNV", "GBA_CHA-CNV")
	AppendSheet(gbaXlsx, finalXlsx, "GBA-variants", "GBA-variants")
	AppendSheet(comXlsx, finalXlsx, "report", "CAH-report")
	AppendSlice2Excel(finalXlsx, "MA-result", maSlice)

	var ma, _ = textUtil.File2MapArray(*ma, "\t", nil)
	var FusionResult = make(map[string]string)
	var avdExtra []map[string]string
	var avdSheetName = "All variants data"
	var avdRaw = simpleUtil.HandleError(finalXlsx.GetRows(avdSheetName)).([][]string)
	var extraTitle = "是否需要验证"
	var title = append(avdRaw[0], extraTitle)
	simpleUtil.CheckErr(
		finalXlsx.SetCellValue(
			avdSheetName,
			simpleUtil.HandleError(
				excelize.CoordinatesToCellName(len(title), 1),
			).(string),
			extraTitle,
		),
	)
	var HBA2NoCheck = map[string]bool{
		"c.369C>G": true,
		"c.377T>C": true,
		"c.427T>C": true,
	}
	var rIdx = len(avdRaw)
	for _, item := range ma {
		var sampleID = item["sample"]
		FusionResult[sampleID] = item["Fusion_result"]
		if item["cHGVS"] != "-" {
			item["SampleID"] = sampleID
			item["A.Depth"] = item["Ad"]
			item["A.Ratio"] = item["Ar"]
			avdExtra = append(avdExtra, item)
			rIdx++
			for j := range title {
				var value, ok = item[title[j]]
				if ok {
					if item["Gene Symbol"] == "HBA2" && HBA2NoCheck[item["cHGVS"]] {
						item[extraTitle] = ""
					} else {
						item[extraTitle] = "验证"
					}
					simpleUtil.CheckErr(
						finalXlsx.SetCellValue(
							avdSheetName,
							simpleUtil.HandleError(
								excelize.CoordinatesToCellName(j+1, rIdx),
							).(string),
							value,
						),
					)
				}
			}
		}
	}
	simpleUtil.CheckErr(finalXlsx.SaveAs(*final + ".OE.xlsx"))
}

func AppendSheet(old, new *excelize.File, oldName, newName string) {
	AppendSlice2Excel(new, newName, simpleUtil.HandleError(old.GetRows(oldName)).([][]string))
}
func AppendSlice2Excel(file *excelize.File, sheetName string, slice [][]string) {
	file.NewSheet(sheetName)
	for i, row := range slice {
		var axis = simpleUtil.HandleError(excelize.CoordinatesToCellName(1, i+1)).(string)
		simpleUtil.CheckErr(file.SetSheetRow(sheetName, axis, &row))
	}
}
