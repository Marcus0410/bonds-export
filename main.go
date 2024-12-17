package main

import (
	"bufio"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

type Allocation struct {
	isin, currency, backOfficeComments, clientName, brokerId, bAndD, feeCurrency string
	qty, infernoNr, smid, book, financeQty                                       int
	tradeDate, valueDate                                                         time.Time
	commitmentFee, price                                                         float64
}

func main() {
	// read input file
	inputFile := getInputFilePath()
	allocations, rullAllocations, tempAllocations, inputPerson, deal, projectId := readInput(inputFile)

	err := writeTradeUpload(allocations, rullAllocations, tempAllocations)
	if err != nil {
		fmt.Println(err)
	}

	err = writeFinance(allocations, inputPerson, deal, projectId)
	if err != nil {
		fmt.Println(err)
	}

	fmt.Println("Output filene har blitt produsert!\nDu kan avslutte programmet.")
	reader := bufio.NewReader(os.Stdin)
	_, err = reader.ReadString('\n')
	if err != nil {
		fmt.Println(err)
	}
}

// get latest modified file in input directory
func getInputFilePath() string {
	inputDir := "./input"

	var inputFileModTime time.Time // modification time of current file
	var inputFilePath string       // path of current file

	// loop through files in input directory
	err := filepath.Walk(inputDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		// if this is a .xlsx file
		if !info.IsDir() && filepath.Ext(path) == ".xlsx" && filepath.Base(path)[0] != '~' {

			// if this file was modified later, this is the current latest modified file
			if info.ModTime().After(inputFileModTime) {
				inputFileModTime = info.ModTime()
				inputFilePath = path
			}
		}

		return nil
	})
	if err != nil {
		fmt.Println(err)
	}

	if inputFilePath == "" {
		fmt.Println("Fant ingen fil i input-mappen.")
	}
	return inputFilePath
}

// read the data in the input file
// returns allocations, rull allocations, temp allocations,
func readInput(inputFilePath string) ([]Allocation, []Allocation, []Allocation, string, string, string) {
	allocations := []Allocation{}
	rullAllocations := []Allocation{}
	tempAllocations := []Allocation{}

	// column indicies, this is used for getting allocation values from Excel
	colInferno := 0
	colInvestor := 1
	colBandD := 2
	colQty := 3
	colRullQty := 4
	colTempQty := 5
	colBroker := 6
	colFee := 7
	colComment := 8
	colFinance := 9
	colBrokerId := 10

	fmt.Println("Leser input-fil med navn:", filepath.Base(inputFilePath))

	file, err := excelize.OpenFile(inputFilePath)
	if err != nil {
		fmt.Println(err)
	}
	defer file.Close()

	sheetName := "Sheet1"

	// get rows
	rows, err := file.GetRows(sheetName)
	if err != nil {
		fmt.Println(err)
	}

	// get corp values
	isin, err := file.GetCellValue(sheetName, "B2")
	rullIsin, err := file.GetCellValue(sheetName, "B3")
	tempIsin, err := file.GetCellValue(sheetName, "B4")
	price, err := file.GetCellValue(sheetName, "B5")
	rullPrice, err := file.GetCellValue(sheetName, "B6")
	tempPrice, err := file.GetCellValue(sheetName, "B7")
	tradeDate, err := file.GetCellValue(sheetName, "B8")
	valueDate, err := file.GetCellValue(sheetName, "B9")
	inputPerson, err := file.GetCellValue(sheetName, "B11")
	deal, err := file.GetCellValue(sheetName, "B12")
	projectId, err := file.GetCellValue(sheetName, "B13")

	if err != nil {
		fmt.Println("Kunne ikke hente en verdi fra Corp-tabellen\n", err)
	}

	// get settlement values
	book, err := file.GetCellValue(sheetName, "E2")
	smid, err := file.GetCellValue(sheetName, "E3")
	rullSmid, err := file.GetCellValue(sheetName, "E4")
	tempSmid, err := file.GetCellValue(sheetName, "E5")
	currency, err := file.GetCellValue(sheetName, "F3")
	rullCurrency, err := file.GetCellValue(sheetName, "F4")
	tempCurrency, err := file.GetCellValue(sheetName, "F5")
	if err != nil {
		fmt.Println("Kunne ikke hente en verdi fra Settlement-tabellen\n", err)
	}

	tradeDateLayout := "1/2/06 15:04"
	valueDateLayout := "01-02-06" // Day.Month.Year

	for rowIndex, row := range rows[16:] {
		var newAlloc Allocation

		// add corp values
		newAlloc.isin = isin
		newAlloc.price, err = strconv.ParseFloat(price, 64)
		if err != nil {
			fmt.Println("Kunne ikke konvertere Price i celle: B5\n", err)
		}
		newAlloc.tradeDate, err = time.Parse(tradeDateLayout, tradeDate)
		if err != nil {
			fmt.Println("Kunne ikke konvertere Trade date i celle: B8\n", err)
		}
		newAlloc.valueDate, err = time.Parse(valueDateLayout, valueDate)
		if err != nil {
			fmt.Println("Kunne ikke konvertere Value date i celle: B9\n", err)
		}
		newAlloc.currency = currency

		// commitment fee currency
		newAlloc.feeCurrency = currency

		// add settlement values
		newAlloc.book, err = strconv.Atoi(book)
		if err != nil {
			fmt.Println("Kunne ikke konvertere Book i celle: E2\n", err)
		}
		newAlloc.smid, err = strconv.Atoi(smid)
		if err != nil {
			fmt.Println("Kunne ikke konvertere SMID i celle: E3\n", err)
		}

		// get values from rows
		for cellIndex, cell := range row {
			// skip empty cells
			if cell == "" {
				continue
			}
			cell = strings.TrimSpace(cell)
			cellName := string(rune(65+cellIndex)) + strconv.Itoa(17+rowIndex) // example: B2

			switch cellIndex {
			case colInferno:
				newAlloc.infernoNr, err = strconv.Atoi(cell)
			case colInvestor:
				newAlloc.clientName = cell
			case colBandD:
				newAlloc.bAndD = strings.TrimSpace(cell)
			case colQty:
				qty := strings.ReplaceAll(cell, ",", "")
				newAlloc.qty, err = strconv.Atoi(qty)
				if err != nil {
					fmt.Println("Kunne ikke konvertere Allocation i celle:", cellName+"\n", err)
				}
			case colRullQty:
				if cell != "" && cell != "0" {
					newRullAlloc := newAlloc
					newRullAlloc.qty, err = strconv.Atoi(strings.ReplaceAll(cell, ",", ""))
					if err != nil {
						fmt.Println("Kunne ikke konvertere rull allocation i celle:", cellName+"\n", err)
					}
					newRullAlloc.isin = rullIsin
					newRullAlloc.smid, err = strconv.Atoi(rullSmid)
					if err != nil {
						fmt.Println("Kunne ikke konvertere Rull SMID\n", err)
					}
					newRullAlloc.price, err = strconv.ParseFloat(rullPrice, 64)
					if err != nil {
						fmt.Println("Kunne ikke konvertere Rull price\n", err)
					}
					newRullAlloc.currency = rullCurrency

					rullAllocations = append(rullAllocations, newRullAlloc)
				}
			case colTempQty:
				if cell != "" && cell != "0" {
					newTempAlloc := newAlloc
					newTempAlloc.qty, err = strconv.Atoi(strings.ReplaceAll(cell, ",", ""))
					if err != nil {
						fmt.Println("Kunne ikke konvertere temp allocation i celle:", cellName+"\n", err)
					}
					newTempAlloc.isin = tempIsin
					newTempAlloc.smid, err = strconv.Atoi(tempSmid)
					if err != nil {
						fmt.Println("Kunne ikke konvertere Temp SMID\n", err)
					}
					newTempAlloc.price, err = strconv.ParseFloat(tempPrice, 64)
					if err != nil {
						fmt.Println("Kunne ikke konvertere Rull price\n", err)
					}
					newTempAlloc.currency = tempCurrency

					tempAllocations = append(tempAllocations, newTempAlloc)
				}
			case colBroker:
				continue
			case colFee:
				if cell != "" {
					newAlloc.commitmentFee, err = strconv.ParseFloat(strings.ReplaceAll(cell, ",", ""), 64)
					if err != nil {
						fmt.Println("Kunne ikke konvertere UW fee i celle: "+cellName+"\n", err)
					}
				}
			case colComment:
				newAlloc.backOfficeComments = cell
			case colFinance:
				financeQty := strings.ReplaceAll(cell, ",", "")
				newAlloc.financeQty, err = strconv.Atoi(financeQty)
				if err != nil {
					fmt.Println("Kunne ikke konvertere Finance rapportering i celle: "+cellName+"\n", err)
				}
			case colBrokerId:
				newAlloc.brokerId = cell
			}
		}

		// add the new allocation to allocation list
		allocations = append(allocations, newAlloc)
	}

	// remove all allocations that have 0 qty
	var filteredAllocations []Allocation // allocations without empty elements
	for _, alloc := range allocations {
		if alloc.qty != 0 {
			filteredAllocations = append(filteredAllocations, alloc)
		}
	}
	allocations = filteredAllocations

	return allocations, rullAllocations, tempAllocations, inputPerson, deal, projectId
}

// create the Excel file for Inferno trade upload
func writeTradeUpload(allocations []Allocation, rullAllocations []Allocation, tempAllocations []Allocation) error {
	file := excelize.NewFile()

	allocationSheet := "Allocations"
	rullSheet := "Rull allocations"
	tempSheet := "Temp allocations"

	// create sheets
	file.SetSheetName("Sheet1", allocationSheet)
	file.NewSheet(rullSheet)
	file.NewSheet(tempSheet)

	// add headers
	headers := []string{"Book", "Counterparty", "Primary Security (GUI)",
		"Number of Shares", "Price", "Trade Date", "Value Date", "Settlement Currency", "Back office comments", "Commitment Fee", "Fee Currency"}

	// filter out all allocations that does not have ABG as bAndD
	var filteredAllocations []Allocation
	for _, alloc := range allocations {
		if strings.ToLower(alloc.bAndD) == "abg" {
			filteredAllocations = append(filteredAllocations, alloc)
		}
	}
	allocations = filteredAllocations

	// write column headers
	for i, header := range headers {
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(65+i)), 1), header)
		file.SetCellValue(rullSheet, fmt.Sprintf("%s%d", string(rune(65+i)), 1), header)
		file.SetCellValue(tempSheet, fmt.Sprintf("%s%d", string(rune(65+i)), 1), header)
	}

	// add main allocations
	for i, allocation := range allocations {
		// insert cell values
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(65)), 2+i), allocation.book)
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(66)), 2+i), allocation.infernoNr)
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(67)), 2+i), allocation.smid)
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(68)), 2+i), -allocation.qty)

		//price
		percentPrice := allocation.price / 100
		priceStr := fmt.Sprintf("%f/%d/%s/PC", percentPrice, allocation.smid, allocation.currency)
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(69)), 2+i), priceStr)

		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(70)), 2+i), allocation.tradeDate)
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(71)), 2+i), allocation.valueDate)
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(72)), 2+i), allocation.currency)
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(73)), 2+i), allocation.backOfficeComments)

		// dont insert 0 if no commitment fee, it should be blank instead
		if allocation.commitmentFee != 0 {
			file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(74)), 2+i), allocation.commitmentFee)
		}

		// fee currency
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(75)), 2+i), allocation.feeCurrency)
	}
	// add rull allocations
	for i, allocation := range rullAllocations {
		// insert cell values
		file.SetCellValue(rullSheet, fmt.Sprintf("%s%d", string(rune(65)), 2+i), allocation.book)
		file.SetCellValue(rullSheet, fmt.Sprintf("%s%d", string(rune(66)), 2+i), allocation.infernoNr)
		file.SetCellValue(rullSheet, fmt.Sprintf("%s%d", string(rune(67)), 2+i), allocation.smid)
		file.SetCellValue(rullSheet, fmt.Sprintf("%s%d", string(rune(68)), 2+i), allocation.qty)

		//price
		percentPrice := allocation.price / 100
		priceStr := fmt.Sprintf("%f/%d/%s/PC", percentPrice, allocation.smid, allocation.currency)
		file.SetCellValue(rullSheet, fmt.Sprintf("%s%d", string(rune(69)), 2+i), priceStr)

		file.SetCellValue(rullSheet, fmt.Sprintf("%s%d", string(rune(70)), 2+i), allocation.tradeDate)
		file.SetCellValue(rullSheet, fmt.Sprintf("%s%d", string(rune(71)), 2+i), allocation.valueDate)
		file.SetCellValue(rullSheet, fmt.Sprintf("%s%d", string(rune(72)), 2+i), allocation.currency)
		file.SetCellValue(rullSheet, fmt.Sprintf("%s%d", string(rune(73)), 2+i), allocation.backOfficeComments)

		// dont insert 0 if no commitment fee, it should be blank instead
		if allocation.commitmentFee != 0 {
			file.SetCellValue(rullSheet, fmt.Sprintf("%s%d", string(rune(74)), 2+i), allocation.commitmentFee)
		}

		// fee currency
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(75)), 2+i), allocation.feeCurrency)
	}
	// add temp allocations
	for i, allocation := range tempAllocations {
		// insert cell values
		file.SetCellValue(tempSheet, fmt.Sprintf("%s%d", string(rune(65)), 2+i), allocation.book)
		file.SetCellValue(tempSheet, fmt.Sprintf("%s%d", string(rune(66)), 2+i), allocation.infernoNr)
		file.SetCellValue(tempSheet, fmt.Sprintf("%s%d", string(rune(67)), 2+i), allocation.smid)
		file.SetCellValue(tempSheet, fmt.Sprintf("%s%d", string(rune(68)), 2+i), -allocation.qty)

		//price
		percentPrice := allocation.price / 100
		priceStr := fmt.Sprintf("%f/%d/%s/PC", percentPrice, allocation.smid, allocation.currency)
		file.SetCellValue(tempSheet, fmt.Sprintf("%s%d", string(rune(69)), 2+i), priceStr)

		file.SetCellValue(tempSheet, fmt.Sprintf("%s%d", string(rune(70)), 2+i), allocation.tradeDate)
		file.SetCellValue(tempSheet, fmt.Sprintf("%s%d", string(rune(71)), 2+i), allocation.valueDate)
		file.SetCellValue(tempSheet, fmt.Sprintf("%s%d", string(rune(72)), 2+i), allocation.currency)
		file.SetCellValue(tempSheet, fmt.Sprintf("%s%d", string(rune(73)), 2+i), allocation.backOfficeComments)

		// dont insert 0 if no commitment fee, it should be blank instead
		if allocation.commitmentFee != 0 {
			file.SetCellValue(tempSheet, fmt.Sprintf("%s%d", string(rune(74)), 2+i), allocation.commitmentFee)
		}

		// fee currency
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(75)), 2+i), allocation.feeCurrency)
	}

	// save file
	err := file.SaveAs("output/tradeUpload.xlsx")
	if err != nil {
		fmt.Println(err)
	}

	return err
}

// insert data to Excel file with finance report
func writeFinance(allocations []Allocation, inputPerson string, deal string, projectId string) error {
	file, err := excelize.OpenFile("output/finance.xlsx")
	if err != nil {
		return err
	}
	defer file.Close()

	sheetName := "Input Front"
	// get rows
	rows, err := file.GetRows(sheetName)
	if err != nil {
		return err
	}
	// clear old values
	for rowNr := range rows {
		// skip first row (headers)
		if rowNr == 0 {
			continue
		}

		for col := 4; col <= 9; col++ {
			// get cell name (e.g. "D2")
			cell, _ := excelize.CoordinatesToCellName(col, rowNr+1)
			err = file.SetCellValue(sheetName, cell, "")
			if err != nil {
				return err
			}
		}
	}

	// insert values to input table
	file.SetCellValue(sheetName, "B1", inputPerson)
	file.SetCellValue(sheetName, "B2", deal)
	file.SetCellValue(sheetName, "B3", projectId)
	file.SetCellValue(sheetName, "B4", allocations[0].isin)
	file.SetCellValue(sheetName, "B5", allocations[0].tradeDate.Format("2006.01.02"))
	file.SetCellValue(sheetName, "B6", allocations[0].currency)

	// add allocation rows
	for i, alloc := range allocations {
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", "D", 2+i), alloc.infernoNr)
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", "E", 2+i), alloc.clientName)
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", "F", 2+i), alloc.brokerId)
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", "H", 2+i), float64(alloc.financeQty)*alloc.price)
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", "I", 2+i), alloc.financeQty)
	}

	// update all formulas
	err = file.UpdateLinkedValue()
	if err != nil {
		return err
	}
	file.Save()
	return err
}
