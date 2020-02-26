package main

import xlst "github.com/ivahaev/go-xlsx-templater"

func Report() {
	doc := xlst.New()
	doc.ReadTemplate("docs/templateKP.xlsx")
	ctx := map[string]interface{}{
		"projectName":       "Покупка нового компьютера",
		"name":              "Mi Pro",
		"deliveryCondition": "Доставка с компаний Винни Пух",
		"paymentCondition":  "Отплата через карту Kaspi",
		"warranty":          "1 год",
		"deadline":          "будет готов через неделю 22.02.20",
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
	doc.Render(ctx)
	doc.Save("./report.xlsx")
}
