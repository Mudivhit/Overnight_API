package main

import (
	"encoding/json"
	"fmt"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"os"
	"strconv"
	"strings"
	"time"
)

type Rows struct {
	Rows []Row `json:"Row"`
}

type Row struct {
	TIMESTAMP string  `json:"TIMESTAMP"`
	OPEN      float64 `json:"OPEN"`
	HIGH      float64 `json:"HIGH"`
	LOW       float64 `json:"LOW"`
	CLOSE     float64 `json:"CLOSE"`
	BID       float64 `json:"BID"`
	ASK       float64 `json:"ASK"`
}

func rowMaker(cell *xlsx.Cell, row *xlsx.Row, sheet *xlsx.Sheet){

	row = sheet.AddRow()
	row.SetHeightCM(0.5)
	cell = row.AddCell()
	cell.Value ="DATE"
	cell = row.AddCell()
	cell.Value ="TIME"
	cell = row.AddCell()
	cell.Value ="OPEN"
	cell = row.AddCell()
	cell.Value ="HIGH"
	cell = row.AddCell()
	cell.Value ="LOW"
	cell = row.AddCell()
	cell.Value ="CLOSE"
	cell = row.AddCell()
	cell.Value ="ASK"
	cell = row.AddCell()
	cell.Value ="BID"
}

func rowPop(cell *xlsx.Cell, row *xlsx.Row, sheet *xlsx.Sheet,date time.Time, time time.Time, bid string, ask string, close string, high string, low string, open string ){
	row = sheet.AddRow()
	row.SetHeightCM(0.5)
	cell = row.AddCell()
	cell.SetDate(date)
	cell.SetFormat("D MMM YYYY")
	cell = row.AddCell()
	cell.SetDateTime(time)
	cell.SetFormat("HH:MM AM/PM")
	cell = row.AddCell()
	cell.Value = open
	cell = row.AddCell()
	cell.Value = high
	cell = row.AddCell()
	cell.Value = low
	cell = row.AddCell()
	cell.Value = close
	cell = row.AddCell()
	cell.Value = ask
	cell = row.AddCell()
	cell.Value = bid

}

func timeConv(tstamp string) time.Time{
	var date = strings.Replace(tstamp, "+00:00", "Z",1)

	result, err := time.Parse(time.RFC3339, date)
	if err != nil {
		panic(err)
	}

	return result
}
func main() {
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row  *xlsx.Row
	var cell *xlsx.Cell
	var err error

	jsonFile, err := os.Open("sss.json")
	// if we os.Open returns an error then handle it
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println("Successfully Opened users.json")
	// defer the closing of our jsonFile so that we can parse it later on

	// read our opened jsonFile as a byte array.
	byteValue, _ := ioutil.ReadAll(jsonFile)
	// we initialize our Users array
	var Rows Rows
	// we unmarshal our byteArray which contains our
	json.Unmarshal(byteValue, &Rows)

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}
	rowMaker(cell, row, sheet)

	for i := 0; i < len(Rows.Rows); i++ {
		var date = timeConv(Rows.Rows[i].TIMESTAMP)
		var bid = strconv.FormatFloat(Rows.Rows[i].BID, 'f', 6, 32)
		var ask = strconv.FormatFloat(Rows.Rows[i].ASK, 'f', 6, 32)
		var close = strconv.FormatFloat(Rows.Rows[i].CLOSE, 'f', 6, 32)
		var high = strconv.FormatFloat(Rows.Rows[i].HIGH, 'f', 6, 32)
		var low = strconv.FormatFloat(Rows.Rows[i].LOW, 'f', 6, 32)
		var open = strconv.FormatFloat(Rows.Rows[i].OPEN, 'f', 6, 32)
		rowPop(cell , row , sheet , date, date , bid , ask , close , high , low , open)
	}
	err = file.Save("lets_see.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}
	jsonFile.Close()
}
