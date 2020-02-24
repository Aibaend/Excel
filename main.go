package main

import (
	_ "image/png"
)

func main() {
	KPBillReport()
	KPReport()
	//BZReport()
	//TXReport()

}

//func Example()  {
//	f, err := excelize.OpenFile("docs/empty.xlsx")
//	if err != nil {
//		println(err)
//		return
//	}
//	f.DuplicateRow("Лист1",5)
//	f.MergeCell("Лист1","A6","D6")
//
//	f.Save()
//}
