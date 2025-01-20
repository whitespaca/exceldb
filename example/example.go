package main

import (
	"fmt"
	"log"
)

func main() {
	db, err := exceldb.NewExcelDatabase("example.xlsx", "Sheet1")
	if err != nil {
		log.Fatalf("Failed to initialize Excel database: %v", err)
	}

	err = db.Insert(exceldb.Row{"ID": "1", "Name": "Alice", "Age": "25"})
	if err != nil {
		log.Fatalf("Failed to insert data: %v", err)
	}
	fmt.Println("Inserted data successfully!")

	results, err := db.Select(exceldb.Row{"Name": "Alice"})
	if err != nil {
		log.Printf("No matching rows found: %v", err)
	} else {
		fmt.Printf("Found data: %+v\n", results)
	}

	err = db.Update(exceldb.Row{"Name": "Alice"}, exceldb.Row{"Age": "26"})
	if err != nil {
		log.Fatalf("Failed to update data: %v", err)
	}
	fmt.Println("Updated data successfully!")

	err = db.Delete(exceldb.Row{"ID": "1"})
	if err != nil {
		log.Fatalf("Failed to delete data: %v", err)
	}
	fmt.Println("Deleted data successfully!")

	err = db.AddColumn("Country", "Unknown")
	if err != nil {
		log.Fatalf("Failed to add column: %v", err)
	}
	fmt.Println("Added new column successfully!")

	err = db.RemoveColumn("Country")
	if err != nil {
		log.Fatalf("Failed to remove column: %v", err)
	}
	fmt.Println("Removed column successfully!")

	sheetNames, err := db.GetAllSheetNames()
	if err != nil {
		log.Fatalf("Failed to get sheet names: %v", err)
	}
	fmt.Printf("Sheet names: %v\n", sheetNames)

	exists, err := db.IsSheetExists("Sheet1")
	if err != nil {
		log.Fatalf("Failed to check sheet existence: %v", err)
	}
	if exists {
		fmt.Println("Sheet1 exists!")
	} else {
		fmt.Println("Sheet1 does not exist.")
	}
}
