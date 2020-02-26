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
	f.SetCellValue(sheet, "E4", "Number of BZ") // номер бланка заказа
	f.SetCellValue(sheet, "C6", "Aim of BZ")    //Назначение
	f.SetCellValue(sheet, "L6", 4)              //Листов Заказа
	f.SetCellValue(sheet, "C8", "Due date")     //Дата постпуление заказа
	f.SetCellValue(sheet, "L8", 8)              //Листов ТХ
	f.SetCellValue(sheet, "C10", "Customer")    // Заказчик
	f.SetCellValue(sheet, "L10", "9999")        // Сумма оплаты
	f.SetCellValue(sheet, "L12", "Форма оплаты")
	f.SetCellValue(sheet, "C14", "Customer info") //ФИО и тел. конт. лица
	f.SetCellValue(sheet, "L14", "Потверждение оплаты")
	f.SetCellValue(sheet, "C16", "Work days") //Дни рабочие по КП либо Договору
	f.SetCellValue(sheet, "L16", "Договор")
	f.SetCellValue(sheet, "C17", "Address")    //Address
	f.SetCellValue(sheet, "L17", "24.02.2020") //Дата сдачи
	starter := 21
	var array []int
	array = append(array, 7, 8, 9)
	///configure table
	for j := 0; j < len(array)-1; j++ {
		f.DuplicateRow(sheet, starter)
		f.MergeCell(sheet, "B22", "D22")
		f.MergeCell(sheet, "E22", "F22")
		f.SetCellValue(sheet, "B22", j)
	}

	for i := range array {
		f.SetCellValue(sheet, "A"+strconv.Itoa(starter+i), i+1)
		f.SetCellValue(sheet, "B"+strconv.Itoa(starter+i), "Name")         //Наименования
		f.SetCellValue(sheet, "E"+strconv.Itoa(starter+i), "Number of TX") //Номер ТХ
		f.SetCellValue(sheet, "G"+strconv.Itoa(starter+i), "Weight")       //Ширина
		f.SetCellValue(sheet, "H"+strconv.Itoa(starter+i), "Depth")        //Глубина
		f.SetCellValue(sheet, "I"+strconv.Itoa(starter+i), "Height")       //Высота
		f.SetCellValue(sheet, "J"+strconv.Itoa(starter+i), 2)              //Количество
		f.SetCellValue(sheet, "K"+strconv.Itoa(starter+i), 2000)           // Цена
		f.SetCellValue(sheet, "L"+strconv.Itoa(starter+i), 1000)           //Сумма

	}

	f.SetCellValue(sheet, "J"+strconv.Itoa(22+len(array)), "Material Info")   //Материал основы/толщины
	f.SetCellValue(sheet, "J"+strconv.Itoa(23+len(array)), "Material Info")   //Материал отделки/вид/цвет
	f.SetCellValue(sheet, "J"+strconv.Itoa(24+len(array)), "Material Info")   //Покрытие
	f.SetCellValue(sheet, "J"+strconv.Itoa(25+len(array)), "Material Info")   //Кромка/толщина/ цвет
	f.SetCellValue(sheet, "I"+strconv.Itoa(26+len(array)), "Material Info")   //Цвет, вид, код ручки
	f.SetCellValue(sheet, "E"+strconv.Itoa(27+len(array)), "Additional Info") //Доп., информация:

	f.SaveAs("docs/resultBZ.xlsx")
}
