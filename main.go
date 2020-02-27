package main

import (
	"encoding/json"
	"fmt"
	"github.com/gorilla/mux"
	_ "image/png"
	"io"
	"net/http"
	"os"
	"strconv"
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

	path, err := GenerateReport(response.Data, response.Template)
	if err != nil {
		fmt.Print(err)
		Respond(w, Message(false, "Template error"))
		return
	}

	Openfile, err := os.Open(path)
	defer Openfile.Close() //Close after function return
	if err != nil {
		//File not found, send 404
		http.Error(w, "File not found.", 404)
		return
	}
	//Get the Content-Type of the file
	//Create a buffer to store the header of the file in
	FileHeader := make([]byte, 512)
	//Copy the headers into the FileHeader buffer
	Openfile.Read(FileHeader)
	//Get content type of file
	FileContentType := http.DetectContentType(FileHeader)

	//Get the file size
	FileStat, _ := Openfile.Stat()                     //Get info from file
	FileSize := strconv.FormatInt(FileStat.Size(), 10) //Get file size as a string

	//Send the headers
	w.Header().Set("Content-Disposition", "attachment; filename="+path)
	w.Header().Set("Content-Type", FileContentType)
	w.Header().Set("Content-Length", FileSize)

	//Send the file
	//We read 512 bytes from the file already, so we reset the offset back to 0
	Openfile.Seek(0, 0)
	io.Copy(w, Openfile) //'Copy' the file to the client
	return
}
