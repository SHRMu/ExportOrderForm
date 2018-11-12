package main

import (
	"database/sql"
	"encoding/json"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	_ "github.com/go-sql-driver/mysql"
	"html/template"
	"log"
	"net/http"
	"os"
)

type Configuration struct {
	Username string
	Password string
	Database string
}

//read info from local config.json
func loadConfig() string{
	file, _ := os.Open("config.json")
	defer file.Close()
	decoder := json.NewDecoder(file)
	configuration := Configuration{}
	err := decoder.Decode(&configuration)
	if err != nil {
		fmt.Println("error:", err)
	}
	config := configuration.Username+":"+configuration.Password+"@/"+configuration.Database
	return config
}

func getCustomerInfo(customername string) (string, string){
	config := loadConfig()
	db, err1 := sql.Open("mysql", config)
	if err1 != nil {
		panic(err1.Error())  // Just for example purpose. You should use proper error handling instead of panic
	}
	defer db.Close()
	var (
		address string
		telephone string
	)
	rows, err := db.Query("select address, telephone from customer_info where customername = ?", customername)
	if err != nil {
		log.Fatal(err)
	}
	defer rows.Close()
	for rows.Next() {
		err := rows.Scan(&address, &telephone)
		if err != nil {
			log.Fatal(err)
		}
	}
	err = rows.Err()
	if err != nil {
		log.Fatal(err)
	}
	return address, telephone
}

func index(w http.ResponseWriter, r *http.Request) {
	var filePath = "./OrderForm.xlsx"
	if r.Method == "GET" {
		t, _ := template.ParseFiles("fillingOrderForm.html")
		t.Execute(w, nil)
	} else {
		r.ParseForm()
		xlsx, err1 := excelize.OpenFile(filePath)
		if err1 != nil {
			fmt.Println(err1)
			return
		}
		xlsx.SetCellValue("Sheet1", "A48", r.Form["applicant"][0])
		customerName := r.Form["customer"][0]
		address, telephone := getCustomerInfo(customerName)
		xlsx.SetCellValue("Sheet1", "E9",customerName)
		fmt.Println(address,"-------",telephone)
		xlsx.SetCellValue("Sheet1", "E11",address)
		xlsx.SetCellValue("Sheet1", "E13",telephone)
		err2 := xlsx.SaveAs(filePath)
		if err2 != nil {
			fmt.Println(err2)
		}
		fmt.Fprintln(w, "Sucess!")
	}
}

func main() {
	http.HandleFunc("/", index)              // set url path
	err := http.ListenAndServe(":9090", nil) //set port to be listened
	if err != nil {
		log.Fatal("ListenAndServe: ", err)
	}
}
