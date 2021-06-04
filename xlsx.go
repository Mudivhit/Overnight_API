package xlsx

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"strings"
	"time"
	"github.com/tealeg/xlsx"
)

type PastMarketPrices struct {
	Row []Row `json:"Row"`
}

type Row struct {
	Open      float64 `json:"OPEN"`
	High      float64 `json:"HIGH"`
	Low       float64 `json:"LOW"`
	Close     float64 `json:"CLOSE"`
	Bid       float64 `json:"BID"`
	Ask       float64 `json:"ASK"`
	Timestamp string  `json:"TIMESTAMP"`
}

func timeConv(tstamp string) time.Time {
	var date = strings.Replace(tstamp, "+00:00", "Z", 1)

	result, err := time.Parse(time.RFC3339, date)
	if err != nil {
		panic(err)
	}

	return result
}

func main() {
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var firstRow, nextRows *xlsx.Row
	var cell *xlsx.Cell
	var err error

	jsonDataFromFile, err := ioutil.ReadFile("./GetIntradayTimeSeries_Response_5.json")

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")

	if err != nil {
		fmt.Println(err.Error())
	}

	var jsonData map[string]interface{}
	err = json.Unmarshal([]byte(jsonDataFromFile), &jsonData)

	b := make([]byte, 10)
	b, err = json.Marshal(jsonData["GetIntradayTimeSeries_Response_5"])
	err = json.Unmarshal(b, &jsonData)

	ls := make([]byte, 10)
	ls, err = json.Marshal(jsonData["Row"])
	n := make([]interface{}, 10)
	err = json.Unmarshal(ls, &n)

	firstRow = sheet.AddRow()
	firstRow.SetHeightCM(0.5)

	cell = firstRow.AddCell()
	cell.Value = "Date"

	cell = firstRow.AddCell()
	cell.Value = "Time"

	cell = firstRow.AddCell()
	cell.Value = "Open"

	cell = firstRow.AddCell()
	cell.Value = "High"

	cell = firstRow.AddCell()
	cell.Value = "Low"

	cell = firstRow.AddCell()
	cell.Value = "Close"

	cell = firstRow.AddCell()
	cell.Value = "Bid"

	cell = firstRow.AddCell()
	cell.Value = "Ask"

	for _, val := range n {
		var row Row

		nextRows = sheet.AddRow()
		nextRows.SetHeightCM(0.5)

		item := make([]byte, 10)
		item, err = json.Marshal(val)
		json.Unmarshal(item, &row)

		var date = timeConv(row.Timestamp)

		cell = nextRows.AddCell()
		cell.SetDate(date)
		cell.SetFormat("D MMM YYYY")
		cell = nextRows.AddCell()
		cell.SetDateTime(date)
		cell.SetFormat("HH:MM AM/PM")

		cell = nextRows.AddCell()
		cell.Value = fmt.Sprintf("%f", row.Open)

		cell = nextRows.AddCell()
		cell.Value = fmt.Sprintf("%f", row.High)

		cell = nextRows.AddCell()
		cell.Value = fmt.Sprintf("%f", row.Low)

		cell = nextRows.AddCell()
		cell.Value = fmt.Sprintf("%f", row.Close)

		cell = nextRows.AddCell()
		cell.Value = fmt.Sprintf("%f", row.Bid)

		cell = nextRows.AddCell()
		cell.Value = fmt.Sprintf("%f", row.Ask)

		// cell = nextRows.AddCell()
		// cell.Value = row.Timestamp

	}

	err = file.Save("GetIntradayTimeSeries_excel.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}

}
