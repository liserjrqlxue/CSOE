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
		"worKDir",
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
		"gba.PE_SE.xlsx, default is -workDir/CAH_GBA/$(basename -workDir)_CAH_GBA.gba.PE_SE.xlsx",
	)
	com = flag.String(
		"com",
		"",
		"com_snp.xlsx, default is -workDir/CAH_GBA/$(basename -workDir)_CAH_GBA.com_snp.xlsx",
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
			*gba = filepath.Join(*workDir, "CAH_GBA", baseDir+"_CAH_GBA.gba.PE_SE.xlsx")
		}
		if *com == "" {
			*com = filepath.Join(*workDir, "CAH_GBA", baseDir+"_CAH_GBA.com_snp.xlsx")
		}
	}
	if *ma == "" || *cnv == "" || *gba == "" || *com == "" {
		fmt.Println("-workDir or -ma/-cnv/-gba/-com are required!")
		os.Exit(1)
	}
	if *final == "" {
		fmt.Println("-final is required!")
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
