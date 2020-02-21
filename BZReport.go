package main

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

type Order struct {
	Aim          string `json:"aim"`
	AcceptDate   string `json:"accept_date"`
	СustomerName string `json:"сustomer_name"`
}

func BZReport() {
	f, err := excelize.OpenFile("docs/templateBZ.xlsx")
	if err != nil {
		println(err.Error())
		return
	}
	sheet = "БЗ2"
	f.SetCellValue(sheet, "E4", "Number of BZ")   // номер бланка заказа
	f.SetCellValue(sheet, "C6", "Aim of BZ")      //Назначение
	f.SetCellValue(sheet, "C8", "Due date")       //Дата постпуление заказа
	f.SetCellValue(sheet, "C20", "Customer")      // Заказчик
	f.SetCellValue(sheet, "C14", "Customer info") //ФИО и тел. конт. лица
	f.SetCellValue(sheet, "C16", "Work days")     //Дни рабочие по КП либо Договору
	f.SetCellValue(sheet, "C17", "Address")       //Address

	f.SetCellValue(sheet, "J26", "Material Info")   //Материал основы/толщины
	f.SetCellValue(sheet, "J27", "Material Info")   //Материал отделки/вид/цвет
	f.SetCellValue(sheet, "J28", "Material Info")   //Покрытие
	f.SetCellValue(sheet, "J29", "Material Info")   //Кромка/толщина/ цвет
	f.SetCellValue(sheet, "I30", "Material Info")   //Цвет, вид, код ручки
	f.SetCellValue(sheet, "E31", "Additional Info") //Доп., информация:

	starter := 21

	var array []int
	for i := range array {
		f.SetCellValue(sheet, "A"+strconv.Itoa(starter+i), i+1)
		f.SetCellValue(sheet, "B"+strconv.Itoa(starter+i), "Name")         //Наименования
		f.SetCellValue(sheet, "E"+strconv.Itoa(starter+i), "Number of TX") //Номер ТХ
		f.SetCellValue(sheet, "G"+strconv.Itoa(starter+i), "Weight")       //Ширина
		f.SetCellValue(sheet, "H"+strconv.Itoa(starter+i), "Depth")        //Глубина
		f.SetCellValue(sheet, "I"+strconv.Itoa(starter+i), "Height")       //Высота
		f.SetCellValue(sheet, "J"+strconv.Itoa(starter+i), "Count")        //Количество
		f.SetCellValue(sheet, "K"+strconv.Itoa(starter+i), "Price")        // Цена
		f.SetCellValue(sheet, "L"+strconv.Itoa(starter+i), "SUM")          //Сумма
		f.InsertRow(sheet, starter+i)
	}

	f.SaveAs("docs/resultBZ.xlsx")
}
