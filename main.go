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
	isin, currency, backOfficeComments, clientName, brokerId, tempIsin, rullIsin string
	qty, infernoNr, smid, book, financeQty, rullQty, tempQty                     int
	tradeDate, valueDate                                                         time.Time
	commitmentFee, price, rullPrice                                              float64
}

func main() {
	// read input file
	inputFile := getInputFilePath()
	allocations, inputPerson, deal, projectId := readInput(inputFile)

	err := writeTradeUpload(allocations)
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
func readInput(inputFilePath string) ([]Allocation, string, string, string) {
	allocations := []Allocation{}

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
	tradeDate, err := file.GetCellValue(sheetName, "B7")
	if err != nil {
		log.Fatal("Kunne ikke hente Trade date\n", err)
	}
	valueDate, err := file.GetCellValue(sheetName, "B8")
	if err != nil {
		log.Fatal("Kunne ikke hente Value date\n", err)
	}
	currency, err := file.GetCellValue(sheetName, "B9")
	if err != nil {
		log.Fatal("Kunne ikke hente Settlement currency\n", err)
	}
	inputPerson, err := file.GetCellValue(sheetName, "B10")
	if err != nil {
		log.Fatal("Kunne ikke hente Input person\n", err)
	}
	deal, err := file.GetCellValue(sheetName, "B11")
	if err != nil {
		log.Fatal("Kunne ikke hente Deal\n", err)
	}
	projectId, err := file.GetCellValue(sheetName, "B12")
	if err != nil {
		log.Fatal("Kunne ikke hente ProjectID\n", err)
	}

	// get settlement values
	book, err := file.GetCellValue(sheetName, "E2")
	if err != nil {
		log.Fatal("Kunne ikke hente Book\n", err)
	}

	timeLayout := "01-02-06" // Day.Month.Year

	// loop through allocations
	for _, row := range rows[14:] {
		var newAlloc Allocation

		// corp values
		newAlloc.isin = isin
		newAlloc.tempIsin = tempIsin
		newAlloc.rullIsin = rullIsin
		newAlloc.price, err = strconv.ParseFloat(price, 64)
		if err != nil {
			log.Fatal("Kunne ikke konvertere Price\n", err)
		}
		newAlloc.rullPrice, err = strconv.ParseFloat(rullPrice, 64)
		if err != nil {
			log.Fatal("Kunne ikke konvertere Rull price\n", err)
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

		newAlloc.clientName = row[0]
		newAlloc.qty, err = strconv.Atoi(strings.ReplaceAll(row[1], ",", ""))
		if err != nil {
			log.Fatal("Kunne ikke konvertere allocation quantity\n", err)
		}

		newAlloc.rullQty, err = strconv.Atoi(row[2])
		if err != nil {
			log.Fatal("Kunne ikke konvertere rull allocation\n", err)
		}

		newAlloc.tempQty, err = strconv.Atoi(row[3])
		if err != nil {
			log.Fatal("Kunne ikke konvertere temp allocation\n", err)
		}

		newAlloc.infernoNr, err = strconv.Atoi(row[4])
		if err != nil {
			log.Fatal("Kunne ikke konvertere Inferno nr\n", err)
		}

		newAlloc.brokerId = row[5]

		// if UW FEE is not empty
		if strings.TrimSpace(row[6]) != "" {
			newAlloc.commitmentFee, err = strconv.ParseFloat(row[6], 64)
			if err != nil {
				log.Fatal("Kunne ikke konvertere UW FEE\n", err)
			}
		}

		newAlloc.backOfficeComments = row[7]
		newAlloc.financeQty, err = strconv.Atoi(strings.ReplaceAll(row[8], ",", ""))
		if err != nil {
			log.Fatal("Kunne ikke konvertere Finance rapportering\n", err)
		}
		allocations = append(allocations, newAlloc)
	}

	return allocations, inputPerson, deal, projectId
}

// create the Excel file for Inferno trade upload
func writeTradeUpload(allocations []Allocation) error {
	file := excelize.NewFile()

	sheetName := "Sheet1"
	// add headers
	headers := []string{"Book", "Counterparty", "Primary Security (GUI)",
		"Number of Shares", "Price", "Trade Date", "Value Date", "Settlement Currency", "Back office comments", "Commitment Fee"}
	//

	// write column headers
	for i, header := range headers {
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", string(rune(65+i)), 1), header)
	}

	// add rows
	for i, allocation := range allocations {
		// insert cell values
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", string(rune(65)), 2+i), allocation.book)
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", string(rune(66)), 2+i), allocation.infernoNr)
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", string(rune(67)), 2+i), allocation.smid)
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", string(rune(68)), 2+i), allocation.qty)

		//price
		percentPrice := allocation.price / 100
		priceStr := fmt.Sprintf("%f/%d/%s/PC", percentPrice, allocation.smid, allocation.currency)
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", string(rune(69)), 2+i), priceStr)

		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", string(rune(70)), 2+i), allocation.tradeDate)
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", string(rune(71)), 2+i), allocation.valueDate)
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", string(rune(72)), 2+i), allocation.currency)
		file.SetCellValue(sheetName, fmt.Sprintf("%s%d", string(rune(73)), 2+i), allocation.backOfficeComments)

		// dont insert 0 if no commitment fee, it should be blank instead
		if allocation.commitmentFee != 0 {
			file.SetCellValue(sheetName, fmt.Sprintf("%s%d", string(rune(74)), 2+i), allocation.commitmentFee)
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
