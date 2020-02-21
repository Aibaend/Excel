package main

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/tealeg/xlsx"
	"strconv"
	"strings"
)

var sheet = "Лист1"
var company = "{my_company_name}"
var bin = "{BIN}"
var iik = "{iik}"
var kbe = "{kbe}"
var bank = "{BANK}"
var bik = "{BIK}"
var payment_code = "{code_of_payment}"
var operation_number = "{operation_number}"
var address = "{operation_number}"
var city = "{city}"
var agreement = "{agreement}"
var payment = "{payment}"
var paymentword = "{payment_word}"
var executor = "{executor}"

type AgreementInfo struct {
	OwnerCompany    *Company   `json:"owner_company"`
	ClientCompany   *Company   `json:"client_company"`
	Date            string     `json:"date"`
	OperationNumber string     `json:"operation_number"`
	Agreement       string     `json:"agreement"`
	Payment         int        `json:"payment"`
	PaymentWord     string     `json:"payment_word"`
	Executor        string     `json:"executor"`
	Products        []*Product `json:"products"`
}

type Company struct {
	Name    string
	BIN     string `json:"bin"`
	City    string `json:"city"`
	IIK     string `json:"iik"`
	Bank    string `json:"bank"`
	BIK     string `json:"bik"`
	KBE     string `json:"kbe"`
	Address string `json:"address"`
	IIN     string `json:"iin,omitempty"`
}

type Image struct {
	ID   uint   `json:"id"`
	Link string `json:"link"`
}
type Volume struct {
	Height  int    `json:"height"`
	Width   int    `json:"width"`
	Depth   int    `json:"depth"`
	Measure string `json:"measure"`
}

type Product struct {
	Name        string  `json:"name"`
	Description string  `json:"description"`
	Image       *Image  `json:"image"`
	Count       int     `json:"count"`
	Measure     string  `json:"measure"`
	Price       int     `json:"price"`
	TotalPrice  int     `json:"total_price,omitempty"`
	Volume      *Volume `json:"volume,omitempty"`
}

func KPBillReport() {
	excelFileName := "docs/template2.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		return
	}
	client := &Company{"MUIT", "This is BIN", "Almaty", "THIS is IIK", "CASPKZ",
		"This is BIK", "This is KBE", "microdistrict Mamyr", "9707074010"}
	owner := &Company{"KIT SYSTEMS", "This is BIN", "Almaty", "THIS is IIK",
		"CASPKZ", "This is BIK", "This is KBE", "microdistrict Mamyr", "9707074010"}
	product := &Product{Name: "HDMI", Count: 1, Measure: "штук", Price: 1000, TotalPrice: 1000}
	products := []*Product{}
	products = append(products, product)
	products = append(products, product)

	agreement := AgreementInfo{Date: "1 сентября 2020", Payment: 1000, Executor: "Andrey", OperationNumber: "XXX13", Agreement: "XXX123", PaymentWord: "тысячи теньге",
		OwnerCompany: owner, ClientCompany: client}
	agreement.Products = append(agreement.Products, product)

	for _, sheet := range xlFile.Sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				if strings.Contains(cell.String(), company) {
					result := strings.Replace(cell.String(), company, agreement.ClientCompany.Name, 1)
					cell.Value = result
				}
				if strings.Contains(cell.String(), bin) {
					result := strings.Replace(cell.String(), bin, agreement.ClientCompany.BIN, 1)
					cell.Value = result
				}
				if strings.Contains(cell.String(), iik) {
					result := strings.Replace(cell.String(), iik, agreement.ClientCompany.IIK, 1)
					cell.Value = result
				}
				if strings.Contains(cell.String(), kbe) {
					result := strings.Replace(cell.String(), kbe, agreement.ClientCompany.KBE, 1)
					cell.Value = result
				}

			}
		}
	}

	xlFile.Save("docs/result.xlsx")
	//B27
	f, err := excelize.OpenFile("docs/result.xlsx")
	if err != nil {
		println(err.Error())
		return
	}

	y := 27
	for i := range products {
		f.SetCellValue(sheet, "B"+strconv.Itoa(y+i), i+1)
		f.SetCellValue(sheet, "D"+strconv.Itoa(y+i), products[i].Name)
		f.SetCellValue(sheet, "T"+strconv.Itoa(y+i), products[i].Count)
		f.SetCellValue(sheet, "X"+strconv.Itoa(y+i), products[i].Measure)
		f.SetCellValue(sheet, "AA"+strconv.Itoa(y+i), products[i].Price)
		f.SetCellValue(sheet, "AG"+strconv.Itoa(y+i), products[i].TotalPrice)
		f.InsertRow(sheet, y+i)
	}

	endRow := 39
	fixedRow := 13
	for i := 0; i < fixedRow-len(products); i++ {
		f.RemoveRow(sheet, endRow-i)
	}

	f.Save()
}
