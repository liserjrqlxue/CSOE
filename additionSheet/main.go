package main

import (
	"flag"
	"path/filepath"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/liserjrqlxue/goUtil/textUtil"
)

// flag
var (
	final = flag.String(
		"final",
		"",
		"final.result.xlsx",
	)
	ma = flag.String(
		"ma",
		filepath.Join("Y159_MA_update", "MA_update.xls"),
		"MA result",
	)
	cnv = flag.String(
		"cnv",
		filepath.Join("CAH_GBA/CNV/CNV_report.reform.xlsx"),
		"CAH_GBA CNV result",
	)
	gba = flag.String(
		"gba",
		"",
		"gba.PE_SE.xlsx",
	)
	com = flag.String(
		"com",
		"",
		"com_snp.xlsx",
	)
)

func main() {
	flag.Parse()
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
