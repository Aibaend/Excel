package main

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

func TXReport() {
	f, err := excelize.OpenFile("docs/templateTX.xlsx")
	if err != nil {
		println(err.Error())
		return
	}
	sheet = "(ТХ)"

	f.SetCellValue(sheet, "C6", "Заказчик")
	//f.SetCellValue(sheet,"С8","ФИО и тел. конт. лица:")
	f.SetCellValue(sheet, "G8", "Назначение")
	f.SetCellValue(sheet, "E10", "Наименование изделия")
	//f.SetCellValue(sheet,"C6","Заказчик")
	//f.SetCellValue(sheet,"C6","Заказчик")
	t := []string{"g", "h", "i"}

	starter := 15
	second_starter := 18 + len(t) - 1
	//settings
	for i := 0; i < len(t)-1; i++ {
		f.DuplicateRow(sheet, starter)

	}
	for j := 0; j < len(t)-1; j++ {
		f.DuplicateRow(sheet, second_starter)
	}
	for i := range t {
		f.SetCellValue(sheet, "A"+strconv.Itoa(starter+i), i+1)
		f.SetCellValue(sheet, "B"+strconv.Itoa(starter+i), "Наименование")
		f.SetCellValue(sheet, "F"+strconv.Itoa(starter+i), "Описание")
		f.SetCellValue(sheet, "A"+strconv.Itoa(second_starter+i), "Наименование детали")
		f.SetCellValue(sheet, "C"+strconv.Itoa(second_starter+i), "Материал")
		f.SetCellValue(sheet, "D"+strconv.Itoa(second_starter+i), "Код-цвет")
		f.SetCellValue(sheet, "E"+strconv.Itoa(second_starter+i), "Толщина")
		f.SetCellValue(sheet, "C"+strconv.Itoa(second_starter+i), "Цвет- код кромки")
		f.SetCellValue(sheet, "G"+strconv.Itoa(second_starter+i), "Материал")
	}

	f.SaveAs("docs/TXresult.xlsx")
}
