package main

import (
	"encoding/json"
	"fmt"
	xlst "github.com/ivahaev/go-xlsx-templater"
)

func GenerateReport(context map[string]interface{}, template string) error {
	doc := xlst.New()
	err := doc.ReadTemplate("docs/" + template)
	if err != nil {
		return err
	}
	doc.Render(context)
	doc.Save("docs/reportTEST.xlsx")

	return err
}

func ReportKP() {
	doc := xlst.New()
	doc.ReadTemplate("docs/templateKP.xlsx")
	ctx := map[string]interface{}{
		"projectName":       "Покупка нового компьютера",
		"name":              "Mi Pro",
		"deliveryCondition": "Доставка с компаний Винни Пух",
		"paymentCondition":  "Отплата через карту Kaspi",
		"warranty":          "1 год",
		"deadline":          " доставка 22.02.20",
		"totalPrice":        "2000000",
		"deliveryPrice":     "5000",
		"generalPrice":      "20005000", //delivery+total
		"products": []map[string]interface{}{
			{
				"id":          1,
				"name":        "MI Pro Notebook",
				"picture":     "picture",
				"description": "Процессор Intel Core I5",
				"weight":      "70 ",
				"height":      "100",
				"depth":       "40",
				"price":       "500 000",
				"count":       1,
				"total":       "500 000",
			},
			{
				"id":          2,
				"name":        "MI Pro Notebook",
				"picture":     "picture",
				"description": "Процессор Intel Core I5",
				"weight":      "70 ",
				"height":      "100",
				"depth":       "40",
				"price":       "500 000",
				"count":       1,
				"total":       "500 000",
			},
			{
				"id":          3,
				"name":        "MI Pro Notebook",
				"picture":     "picture",
				"description": "Процессор Intel Core I5",
				"weight":      "70 ",
				"height":      "100",
				"depth":       "40",
				"price":       "500 000",
				"count":       1,
				"total":       "500 000",
			},
		},
	}
	data, _ := json.Marshal(ctx)
	fmt.Println(string(data))
	doc.Render(ctx)
	doc.Save("./report.xlsx")
}

func ReportKPBill() {
	doc := xlst.New()
	doc.ReadTemplate("docs/templateKPBill.xlsx")
	ctx := map[string]interface{}{
		"beneficiary":     "KIT Systems",
		"BIN":             "970708301791",
		"iik":             "192HVA02",
		"kbe":             "this is KBE",
		"BANK":            "CASPKZ",
		"BIK":             "this is BIK",
		"paymentCode":     "xxx13",
		"operationNumber": "xxx123",
		"date":            "12 сентября 1972",
		"address":         "Micro district Kalkaman",
		"city":            "Apple city",
		"customerBIN":     "740928402507",
		"customerName":    "CalmON",
		"customerAddress": "Koktobe",
		"agreement":       "Number:1892",
		"payment":         "1000000",
		"productsCount":   "2",
		"paymentWord":     "Один миллион",
		"executor":        "SaDU Aibyn",
		"products": []map[string]interface{}{
			{
				"id":      1,
				"name":    "MI Pro Notebook",
				"price":   "500 000",
				"count":   1,
				"total":   "500 000",
				"measure": "штук",
			},
			{
				"id":      1,
				"name":    "MI Pro Notebook",
				"price":   "500 000",
				"count":   1,
				"total":   "500 000",
				"measure": "штук",
			},
		},
	}
	doc.Render(ctx)
	doc.Save("docs/reportKPBill.xlsx")
}

func ReportBZ() {
	doc := xlst.New()
	doc.ReadTemplate("docs/templateBZ.xlsx")
	ctx := map[string]interface{}{
		"number":          "X123X",
		"listOrder":       "2",
		"date":            "22 февраля 2020",
		"listTX":          "3",
		"customer":        "Johny Depp",
		"payment":         "19990",
		"paymentForm":     "наличный",
		"fio":             "Vladimir Putin Vladimirovich",
		"phone":           "8777 111 11 11 ",
		"isPay":           "ДА",
		"workday":         "3 дня",
		"agreementNumber": "1234",
		"address":         "microdistrict  Mamyr",
		"deadline":        "29 февраля 2020",
		"products": []map[string]interface{}{
			{
				"id":     1,
				"name":   "MI Pro Notebook",
				"weight": "70 ",
				"height": "100",
				"depth":  "40",
				"price":  "500 000",
				"count":  1,
				"total":  "500 000",
				"number": "192HVA02",
			},
			{
				"id":     2,
				"name":   "MI Pro Notebook",
				"weight": "70 ",
				"height": "100",
				"depth":  "40",
				"price":  "500 000",
				"count":  1,
				"total":  "500 000",
				"number": "192HVA02",
			},
			{
				"id":     3,
				"name":   "MI Pro Notebook",
				"weight": "70 ",
				"height": "100",
				"depth":  "40",
				"price":  "500 000",
				"count":  1,
				"total":  "500 000",
				"number": "192HVA02",
			},
		},
		"materialWeight":   "89kg",
		"materialView":     "дерево и металь",
		"materialSkin":     "Матовый",
		"materialColor":    "GOLD",
		"materialAddition": "Nothing",
	}

	doc.Render(ctx)
	doc.Save("docs/reportBZ.xlsx")
}
