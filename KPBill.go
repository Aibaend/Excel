package main

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

var sheet = "Лист1"

type KPBill struct {
	Beneficiary     *Company `json:"beneficiary"`
	PaymentCode     string   `json:"payment_code"`
	Date            string   `json:"date"`
	AgreementNumber string   `json:"agreement_number"`
}

type Company struct {
	Name    string `json:"name"`
	BIN     string `json:"bin"`
	IIK     string `json:"iik"`
	KBE     string `json:"kbe"`
	Bank    string `json:"bank"`
	BIK     string `json:"bik"`
	Address string `json:"address"`
	City    string `json:"city"`
}
type Product struct {
	Name       string `json:"name"`
	Count      int    `json:"count"`
	Measure    string `json:"measure"`
	Price      int    `json:"price"`
	TotalPrice int    `json:"total_price"`
	WordPrice  string `json:"word_price"`
}

func KPBillReport() {
	f, err := excelize.OpenFile("docs/templateKPBill.xlsx")
	if err != nil {
		println(err)
		return
	}
	sheet = "Лист1"

	f.SetCellValue(sheet, "B11", "KIT SYSTEMS")   //Бенефициар
	f.SetCellValue(sheet, "B12", "BIN: 32123213") //BIN
	f.SetCellValue(sheet, "B14", "CASPKZ")        //Банк бенефициара
	f.SetCellValue(sheet, "V11", "IIK")           //IIK
	f.SetCellValue(sheet, "V14", "BIK")
	f.SetCellValue(sheet, "AF11", "KBE")
	f.SetCellValue(sheet, "AD14", "CODE 13XXX")
	f.SetCellValue(sheet, "B17", "Счет на оплату N 12314 от 12.01.2019")
	f.SetCellValue(sheet, "F21", "970708301791, KIT SYSTEMS, Тенгиз Товерс 30/1")
	f.SetCellValue(sheet, "F23", "970708301791, KIT SYSTEMS,Алматы, Тенгиз Товерс 30/1")
	f.SetCellValue(sheet, "F25", "5634221")

	products := []*Product{}

	product_one := &Product{Name: " Chair", Count: 2, Measure: "штук", Price: 1000}
	product_two := &Product{Name: " Chair", Count: 2, Measure: "штук", Price: 1000}
	products = append(products, product_one, product_two)
	starter := 28
	for j := 0; j < len(products)-1; j++ {
		f.DuplicateRow(sheet, starter)
		f.MergeCell(sheet, "B29", "C29")
		f.MergeCell(sheet, "D29", "S29")
		f.MergeCell(sheet, "T29", "W29")
		f.MergeCell(sheet, "X29", "Z29")
		f.MergeCell(sheet, "AA29", "AF29")
		f.MergeCell(sheet, "AG29", "AL29")

	}
	var sum int
	for i := range products {
		f.SetCellValue(sheet, "B"+strconv.Itoa(starter+i), i+1)
		f.SetCellValue(sheet, "D"+strconv.Itoa(starter+i), products[i].Name)
		f.SetCellValue(sheet, "T"+strconv.Itoa(starter+i), products[i].Count)
		f.SetCellValue(sheet, "X"+strconv.Itoa(starter+i), products[i].Measure)
		f.SetCellValue(sheet, "AA"+strconv.Itoa(starter+i), products[i].Price)
		f.SetCellValue(sheet, "AG"+strconv.Itoa(starter+i), products[i].Price*products[i].Count)
		sum = sum + products[i].Price*products[i].Count
	}
	f.SetCellValue(sheet, "AG"+strconv.Itoa(29+len(products)), sum)
	f.SetCellValue(sheet, "B"+strconv.Itoa(31+len(products)), "Всего наименований "+strconv.Itoa(len(products))+" на сумму "+strconv.Itoa(sum))
	f.SetCellValue(sheet, "B"+strconv.Itoa(32+len(products)), "Всего к оплате: Столько-то тысяч тенге")
	f.SetCellValue(sheet, "G"+strconv.Itoa(36+len(products)), "Done by Aibaend")
	f.SaveAs("docs/resultKPBill.xlsx")
}
