package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"html/template"
	"log"
	"net/http"
)

func index(w http.ResponseWriter, r *http.Request) {
	var filePath = "./OrderForm.xlsx"
	if r.Method == "GET" {
		t, _ := template.ParseFiles("fillingOrderForm.html")
		t.Execute(w, nil)
	} else {
		r.ParseForm()
		xlsx, err := excelize.OpenFile(filePath)
		if err != nil {
			fmt.Println(err)
			return
		}
		xlsx.SetCellValue("Sheet1", "A48", r.Form["username"])
		error := xlsx.SaveAs(filePath)
		if error != nil {
			fmt.Println(error)
		}
	}
}
func main() {
	http.HandleFunc("/", index) // set url path
	err := http.ListenAndServe(":9090", nil) //set port to be listened
	if err != nil {
		log.Fatal("ListenAndServe: ", err)
	}
}
