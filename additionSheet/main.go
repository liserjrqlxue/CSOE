package main

import (
	"flag"
	"path/filepath"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/liserjrqlxue/goUtil/textUtil"
	"github.com/tealeg/xlsx/v2"
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
	var finalXlsx = simpleUtil.HandleError(xlsx.OpenFile(*final)).(*xlsx.File)
	var cnvXlsx = simpleUtil.HandleError(xlsx.OpenFile(*cnv)).(*xlsx.File)
	var gbaXlsx = simpleUtil.HandleError(xlsx.OpenFile(*gba)).(*xlsx.File)
	var comXlsx = simpleUtil.HandleError(xlsx.OpenFile(*com)).(*xlsx.File)
	simpleUtil.HandleError(finalXlsx.AppendSheet(*cnvXlsx.Sheet["CNV"], "GBA_CHA-CNV"))
	simpleUtil.HandleError(finalXlsx.AppendSheet(*gbaXlsx.Sheet["GBA-variants"], "GBA-variants"))
	simpleUtil.HandleError(finalXlsx.AppendSheet(*comXlsx.Sheet["report"], "CAH-report"))
	var maSheet = simpleUtil.HandleError(finalXlsx.AddSheet("MA")).(*xlsx.Sheet)
	for _, maArray := range maSlice {
		var row = maSheet.AddRow()
		for _, v := range maArray {
			row.AddCell().SetValue(v)
		}

	}
	simpleUtil.CheckErr(finalXlsx.Save(*final + "OS.xlsx"))
}
