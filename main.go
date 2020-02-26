package main

import (
	"encoding/json"
	"fmt"
	"github.com/gorilla/mux"
	_ "image/png"
	"net/http"
	"os"
)

func main() {
	router := mux.NewRouter()
	//ReportKP()
	//ReportKPBill()
	//ReportBZ()
	//KPBillReport()
	//KPReport()
	//BZReport()
	//TXReport()
	router.HandleFunc("/report", handleResponse).Methods("POST")

	port := os.Getenv("PORT")
	if port == "" {
		port = "8078" //localhost
	}
	//handler := LogMiddleware(router)
	fmt.Println("Server is running on port 8078")
	err := http.ListenAndServe(":"+port, router)
	if err != nil {
		fmt.Print(err)
	}

}

func Message(status bool, message string) map[string]interface{} {
	return map[string]interface{}{"status": status, "message": message}
}

func Respond(w http.ResponseWriter, data map[string]interface{}) {
	w.Header().Add("Content-Type", "application/json")
	json.NewEncoder(w).Encode(data)
}

type Response struct {
	Template string
	Data     map[string]interface{}
}

var handleResponse = func(w http.ResponseWriter, r *http.Request) {

	response := &Response{}
	err := json.NewDecoder(r.Body).Decode(&response) //decode the request body into struct and failed if any error occur
	fmt.Print(err)
	if err != nil {
		Respond(w, Message(false, "Invalid request"))
		return
	}

	fmt.Print(response.Data)
	err = GenerateReport(response.Data, response.Template)
	if err != nil {
		fmt.Print(err)
		Respond(w, Message(false, "Template error"))
		return
	}
}
