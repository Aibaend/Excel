package main

import (
	"fmt"
	"github.com/gorilla/mux"
	_ "image/png"
	"net/http"
	"os"
)

func main() {
	router := mux.NewRouter()
	Report()
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

var handleResponse = func(w http.ResponseWriter, r *http.Request) {

}
