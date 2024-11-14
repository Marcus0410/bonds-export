package main

import (
	"bufio"
	"fmt"
	"log"
	"os"
	"time"

	"github.com/xuri/excelize/v2"
)

type Allocation struct {
	isin                       string
	qty, infernoNr, smid, book int
	tradeDate, valueDate       time.Time
}

func main() {
	// read input file
	// allocations, err := readInput("./input/*.xlsx")
	// if err != nil {
	// 	log.Fatal(err)
	// }

	// write output files
	allocations := []Allocation{}
	a1 := Allocation{}
	a1.isin = "ISIN123"
	a1.qty = 123
	a1.infernoNr = 431930
	a1.smid = 583920
	a1.book = 1072

	allocations = append(allocations, a1)
	allocations = append(allocations, a1)
	allocations = append(allocations, a1)
	allocations = append(allocations, a1)
	allocations = append(allocations, a1)

	err := writeTradeUpload(allocations)
	if err != nil {
		log.Fatal(err)
	}

	err = writeFinance(allocations)
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Output filene har blitt produsert!\nTrykk 'enter' for å avslutte programmet...")
	reader := bufio.NewReader(os.Stdin)
	_, err = reader.ReadString('\n')
	if err != nil {
		log.Fatal(err)
	}
}

func readInput(filePath string) ([]Allocation, error) {
	allocations := []Allocation{}

	file, err := excelize.OpenFile(filePath)
	if err != nil {
		return allocations, err
	}
	defer file.Close()

	return allocations, err
}

func writeTradeUpload(allocations []Allocation) error {
	file := excelize.NewFile()

	// add headers
	headers := []string{"Book", "Counterparty", "Primary Security (GUI)",
		"Number of Shares", "Price", "Trade Date", "Value Date", "Settlement Currency"}

	for i, header := range headers {
		file.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(65+i)), 1), header)
	}

	// add rows
	for i, allocation := range allocations {
		file.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(65)), 2+i), allocation.book)
		file.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(66)), 2+i), allocation.infernoNr)
		file.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(67)), 2+i), allocation.smid)
		file.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(68)), 2+i), allocation.qty)
		// file.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(69)), 2+i), allocation.)
	}

	err := file.SaveAs("output/tradeUpload.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	return err
}

func writeFinance(allocations []Allocation) error {
	file := excelize.NewFile()

	// add headers
	headers := []string{"TransID", "Ticker", "Date", "PostingDate", "StaffID", "AccountID",
		"Customer", "ExchangeID", "InstrumentID", "ProductID", "ProjectID", "TradeFX",
		"RevenueTransaction", "ValueTransaction", "Volume", "EnteredBy", "Comment", "RevenueType"}

	for i, header := range headers {
		file.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(65+i)), 1), header)
	}

	// add rows
	// for i, allocation := range allocations {
	// 	file.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(65)), 2+i), allocation.book)
	// 	file.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(66)), 2+i), allocation.infernoNr)
	// 	file.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(67)), 2+i), allocation.smid)
	// 	file.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(68)), 2+i), allocation.qty)
	// 	// file.SetCellValue("Sheet1", fmt.Sprintf("%s%d", string(rune(69)), 2+i), allocation.)
	// }

	err := file.SaveAs("output/finance.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	return err
}
