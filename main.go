package main

import (
	"bufio"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type Allocation struct {
	isin, currency, backOfficeComments, clientName, brokerId, accType string
	qty, infernoNr, smid, book, financeQty                            int
	tradeDate, valueDate                                              time.Time
	commitmentFee, price                                              float64
}

func main() {
	// read input file
	inputFile := getInputFilePath()
	allocations, rullAllocations, tempAllocations, inputPerson, deal, projectId := readInput(inputFile)

	err := writeTradeUpload(allocations, rullAllocations, tempAllocations)
	if err != nil {
		log.Fatal(err)
	}

	err = writeFinance(allocations, inputPerson, deal, projectId)
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Output filene har blitt produsert!\nDu kan avslutte programmet.")
	reader := bufio.NewReader(os.Stdin)
	_, err = reader.ReadString('\n')
	if err != nil {
		log.Fatal(err)
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

		// if this is a file
		if !info.IsDir() && filepath.Ext(path) == ".xlsx" {
			// if this file was modified later, this is the current latest modified file
			if info.ModTime().After(inputFileModTime) {
				inputFileModTime = info.ModTime()
				inputFilePath = path
			}
		}

		return nil
	})
	if err != nil {
		log.Fatal(err)
	}

	if inputFilePath == "" {
		log.Fatal("Fant ingen fil i input-mappen.")
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
	colAccType := 2
	colQty := 3
	colRullQty := 4
	colTempQty := 5
	// colBroker := 6 currently not needed
	colFee := 7
	colComment := 8
	colFinance := 9
	colBrokerId := 10

	fmt.Println("Leser input-fil med navn:", filepath.Base(inputFilePath))

	file, err := excelize.OpenFile(inputFilePath)
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()

	sheetName := "Sheet1"

	// get rows
	rows, err := file.GetRows(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	// get corp values
	isin, err := file.GetCellValue(sheetName, "B2")
	if err != nil {
		log.Fatal("Kunne ikke hente ISIN\n", err)
	}
	tempIsin, err := file.GetCellValue(sheetName, "B3")
	if err != nil {
		log.Fatal("Kunne ikke hente Temp ISIN\n", err)
	}
	rullIsin, err := file.GetCellValue(sheetName, "B4")
	if err != nil {
		log.Fatal("Kunne ikke hente Rull ISIN\n", err)
	}
	price, err := file.GetCellValue(sheetName, "B5")
	if err != nil {
		log.Fatal("Kunne ikke hente price\n", err)
	}
	rullPrice, err := file.GetCellValue(sheetName, "B6")
	if err != nil {
		log.Fatal("Kunne ikke hente Rull price\n", err)
	}
	tempPrice, err := file.GetCellValue(sheetName, "B7")
	if err != nil {
		log.Fatal("Kunne ikke hente Temp price\n", err)
	}
	tradeDate, err := file.GetCellValue(sheetName, "B8")
	if err != nil {
		log.Fatal("Kunne ikke hente Trade date\n", err)
	}
	valueDate, err := file.GetCellValue(sheetName, "B9")
	if err != nil {
		log.Fatal("Kunne ikke hente Value date\n", err)
	}
	currency, err := file.GetCellValue(sheetName, "B10")
	if err != nil {
		log.Fatal("Kunne ikke hente Settlement currency\n", err)
	}
	inputPerson, err := file.GetCellValue(sheetName, "B11")
	if err != nil {
		log.Fatal("Kunne ikke hente Input person\n", err)
	}
	deal, err := file.GetCellValue(sheetName, "B12")
	if err != nil {
		log.Fatal("Kunne ikke hente Deal\n", err)
	}
	projectId, err := file.GetCellValue(sheetName, "B13")
	if err != nil {
		log.Fatal("Kunne ikke hente ProjectID\n", err)
	}

	// get settlement values
	book, err := file.GetCellValue(sheetName, "E2")
	if err != nil {
		log.Fatal("Kunne ikke hente Book\n", err)
	}
	smid, err := file.GetCellValue(sheetName, "E3")
	if err != nil {
		log.Fatal("Kunne ikke hente SMID\n", err)
	}
	rullSmid, err := file.GetCellValue(sheetName, "E4")
	if err != nil {
		log.Fatal("Kunne ikke hente Rull SMID\n", err)
	}
	tempSmid, err := file.GetCellValue(sheetName, "E5")
	if err != nil {
		log.Fatal("Kunne ikke hente Temp SMID\n", err)
	}

	timeLayout := "01-02-06" // Day.Month.Year

	// loop through allocations
	for _, row := range rows[16:] {
		var newAlloc Allocation

		// corp values
		newAlloc.isin = isin
		newAlloc.smid, err = strconv.Atoi(smid)
		if err != nil {
			log.Fatal("Kunne ikke konvertere SMID\n", err)
		}
		newAlloc.price, err = strconv.ParseFloat(price, 64)
		if err != nil {
			log.Fatal("Kunne ikke konvertere Price\n", err)
		}
		newAlloc.tradeDate, err = time.Parse(timeLayout, tradeDate)
		if err != nil {
			log.Fatal("Kunne ikke konvertere Trade date\n", err)
		}
		newAlloc.valueDate, err = time.Parse(timeLayout, valueDate)
		if err != nil {
			log.Fatal("Kunne ikke konvertere Value date\n", err)
		}
		newAlloc.currency = currency

		// settlement values
		newAlloc.book, err = strconv.Atoi(book)
		if err != nil {
			log.Fatal("Kunne ikke konvertere Book\n", err)
		}

		newAlloc.clientName = row[colInvestor]
		newAlloc.qty, err = strconv.Atoi(strings.ReplaceAll(row[colQty], ",", ""))
		if err != nil {
			log.Fatal("Kunne ikke konvertere allocation quantity\n", err)
		}

		newAlloc.infernoNr, err = strconv.Atoi(row[colInferno])
		if err != nil {
			log.Fatal("Kunne ikke konvertere Inferno nr\n", err)
		}

		newAlloc.accType = row[colAccType]

		newAlloc.brokerId = row[colBrokerId]

		// if UW FEE is not empty
		if strings.TrimSpace(row[colFee]) != "" {
			newAlloc.commitmentFee, err = strconv.ParseFloat(row[colFee], 64)
			if err != nil {
				log.Fatal("Kunne ikke konvertere UW FEE\n", err)
			}
		}

		newAlloc.backOfficeComments = row[colComment]
		newAlloc.financeQty, err = strconv.Atoi(strings.ReplaceAll(row[colFinance], ",", ""))
		if err != nil {
			log.Fatal("Kunne ikke konvertere Finance rapportering\n", err)
		}

		// add main allocation to allocations
		allocations = append(allocations, newAlloc)

		// add rull allocation
		if row[colRullQty] != "" && row[colRullQty] != "0" {
			newRullAlloc := newAlloc
			newRullAlloc.qty, err = strconv.Atoi(strings.ReplaceAll(row[colRullQty], ",", ""))
			if err != nil {
				log.Fatal("Kunne ikke konvertere rull allocation\n", err)
			}
			newRullAlloc.isin = rullIsin
			newRullAlloc.smid, err = strconv.Atoi(rullSmid)
			if err != nil {
				log.Fatal("Kunne ikke konvertere Rull SMID\n", err)
			}
			newRullAlloc.price, err = strconv.ParseFloat(rullPrice, 64)
			if err != nil {
				log.Fatal("Kunne ikke konvertere Rull price\n", err)
			}
			rullAllocations = append(rullAllocations, newRullAlloc)
		}

		// add temp allocation
		if row[colTempQty] != "" && row[colTempQty] != "0" {
			newTempAlloc := newAlloc
			newTempAlloc.qty, err = strconv.Atoi(strings.ReplaceAll(row[colTempQty], ",", ""))
			if err != nil {
				log.Fatal("Kunne ikke konvertere temp allocation\n", err)
			}
			newTempAlloc.isin = tempIsin
			newTempAlloc.smid, err = strconv.Atoi(tempSmid)
			if err != nil {
				log.Fatal("Kunne ikke konvertere Temp SMID\n", err)
			}
			newTempAlloc.price, err = strconv.ParseFloat(tempPrice, 64)
			if err != nil {
				log.Fatal("Kunne ikke konvertere Rull price\n", err)
			}
			tempAllocations = append(tempAllocations, newTempAlloc)
		}
	}

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
		"Number of Shares", "Price", "Trade Date", "Value Date", "Settlement Currency", "Back office comments", "Commitment Fee"}

	// write column headers
	for i, header := range headers {
		file.SetCellValue(allocationSheet, fmt.Sprintf("%s%d", string(rune(65+i)), 1), header)
		file.SetCellValue(rullSheet, fmt.Sprintf("%s%d", string(rune(65+i)), 1), header)
		file.SetCellValue(tempSheet, fmt.Sprintf("%s%d", string(rune(65+i)), 1), header)
	}

	// add main allocations
	for i, allocation := range allocations {
		// if account type is Pot, do not include
		if strings.ToLower(strings.TrimSpace(allocation.accType)) == "pot" {
			continue
		}
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
	}

	// save file
	err := file.SaveAs("output/tradeUpload.xlsx")
	if err != nil {
		log.Fatal(err)
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
	file.SetCellValue(sheetName, "B5", allocations[0].tradeDate)
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
