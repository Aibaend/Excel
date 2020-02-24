package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

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
		println(err)
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
	var sum int
	for j := 0; j < len(products)-1; j++ {
		f.DuplicateRow(sheet, starter+j)
	}
	for i := range products {

		f.SetCellValue(sheet, "A"+strconv.Itoa(starter+i), i+1)
		f.SetCellValue(sheet, "B"+strconv.Itoa(starter+i), product.Name+" "+strconv.Itoa(product.Volume.Width)+"X"+
			strconv.Itoa(product.Volume.Height)+"X"+strconv.Itoa(product.Volume.Depth))
		//image paste

		err = f.AddPicture(sheet, "C"+strconv.Itoa(starter+i), "images/table2.png", `{"x_scale": 0.5, "y_scale": 0.45}`)
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
		sum = sum + product.Price*product.Count

	}

	f.SetCellValue(sheet, "J"+strconv.Itoa(16+len(products)), sum)
	fmt.Print(len(products))
	f.SetCellValue(sheet, "J"+strconv.Itoa(17+len(products)), 5000)
	f.SetCellValue(sheet, "J"+strconv.Itoa(18+len(products)), sum+5000) ///sum+ dostavka

	f.SaveAs("docs/resultKP.xlsx")

}
