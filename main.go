package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/ivahaev/go-xlsx-templater"
	"github.com/tealeg/xlsx"
	_ "image/png"
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

func main() {
	KPBillReport()
	KPReport()
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
	}

	endRow := 39
	fixedRow := 13
	for i := 0; i < fixedRow-len(products); i++ {
		f.RemoveRow(sheet, endRow-i)
	}

	f.Save()
}

type Project struct {
	Project           string     `json:"project"`
	Name              string     `json:"name"`
	DeliverConditions string     `json:"conditions"`
	PaymentConditions string     `json:"payment_conditions"`
	Warranty          string     `json:"warranty"`
	Deadline          string     `json:"deadline"`
	Products          []*Product `json:"products"`
}

func KPReport() {
	f, err := excelize.OpenFile("docs/templateKP.xlsx")
	if err != nil {
		println(err.Error())
		return
	}
	sheet = "ЛДСП"
	project := Project{Project: "KIT SYSTEMS", Name: "Hello", DeliverConditions: "упаковка+доставка+монтаж в г. Нур-Султан",
		PaymentConditions: "наличным ", Warranty: "12 месяцев", Deadline: "р.д. После оплаты и всех согласований."}
	product := &Product{Name: "Стол", Description: "Этот стол для того чтобы спасти тебя:)", Volume: &Volume{Width: 250, Height: 2260, Depth: 350, Measure: "cm"},
		Price: 1000000, Count: 10}
	products := []*Product{}
	products = append(products, product)
	products = append(products, product)
	fmt.Print(len(products))

	f.SetCellValue(sheet, "C4", project.Project)
	f.SetCellValue(sheet, "C5", project.Name)
	f.SetCellValue(sheet, "C6", project.DeliverConditions)
	f.SetCellValue(sheet, "C7", project.PaymentConditions)
	f.SetCellValue(sheet, "C8", project.Warranty)
	f.SetCellValue(sheet, "C10", project.Deadline)

	starter := 15
	for i := range products {

		f.SetCellValue(sheet, "A"+strconv.Itoa(starter+i), i+1)
		f.SetCellValue(sheet, "B"+strconv.Itoa(starter+i), product.Name+" "+strconv.Itoa(product.Volume.Width)+"X"+
			strconv.Itoa(product.Volume.Height)+"X"+strconv.Itoa(product.Volume.Depth))
		//image paste
		err = f.AddPicture(sheet, "C"+strconv.Itoa(starter+i), "images/table2.png", `{"x_scale": 0.5, "y_scale": 0.5}`)
		if err != nil {
			fmt.Println(err)
		}
		f.SetCellValue(sheet, "D"+strconv.Itoa(starter+i), product.Description)
		f.SetCellValue(sheet, "E"+strconv.Itoa(starter+i), product.Volume.Width)
		f.SetCellValue(sheet, "F"+strconv.Itoa(starter+i), product.Volume.Depth)
		f.SetCellValue(sheet, "G"+strconv.Itoa(starter+i), product.Volume.Height)
		f.SetCellValue(sheet, "H"+strconv.Itoa(starter+i), product.Price)
		f.SetCellValue(sheet, "I"+strconv.Itoa(starter+i), product.Count)
		f.SetCellValue(sheet, "J"+strconv.Itoa(starter+i), product.Price*product.Count)

	}

	f.SaveAs("docs/resultKP.xlsx")
}
