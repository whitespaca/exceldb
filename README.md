# ExcelDB-go

**ExcelDB** is a lightweight Go package that provides a simple interface for reading, writing, and manipulating Excel files (`.xlsx`) as if they were a basic database. It supports CRUD operations, dynamic column handling, and sheet management using the powerful [excelize](https://github.com/xuri/excelize) library.

---

## 📦 Features

- 🔍 Query (Select) rows based on conditions  
- ➕ Insert new rows  
- ✏️ Update rows matching a query  
- ❌ Delete rows based on criteria  
- ➕➖ Add/Remove columns  
- 📄 Manage multiple sheets  
- 💾 Automatically creates an Excel file if it doesn’t exist

---

## 🛠️ Installation

```bash
go get github.com/whitespaca/exceldb
```

(Replace with your actual module path)

Also install the dependency:

```bash
go get github.com/xuri/excelize/v2
```

---

## 🧪 Usage

```go
package main

import (
	"fmt"
	"log"
    
	"github.com/whitespaca/exceldb"
)

func main() {
	db, err := exceldb.NewExcelDatabase("data.xlsx", "Users")
	if err != nil {
		log.Fatal(err)
	}

	// Insert data
	err = db.Insert(exceldb.Row{"Name": "Alice", "Age": 30})
	if err != nil {
		log.Fatal(err)
	}

	// Query data
	rows, err := db.Select(exceldb.Row{"Name": "Alice"})
	if err != nil {
		log.Println("Not found")
	} else {
		fmt.Println("Found:", rows)
	}

	// Update data
	err = db.Update(exceldb.Row{"Name": "Alice"}, exceldb.Row{"Age": 31})
	if err != nil {
		log.Fatal(err)
	}

	// Delete data
	err = db.Delete(exceldb.Row{"Name": "Alice"})
	if err != nil {
		log.Fatal(err)
	}
}
```

---

## 📚 API Reference

### Initialization

```go
db, err := exceldb.NewExcelDatabase(filePath string, sheetName string)
```

---

### Insert

```go
err := db.Insert(Row{"Name": "Bob", "Age": 25})
```

---

### Select

```go
rows, err := db.Select(Row{"Name": "Bob"})
```

---

### Update

```go
err := db.Update(Row{"Name": "Bob"}, Row{"Age": 26})
```

---

### Delete

```go
err := db.Delete(Row{"Name": "Bob"})
```

---

### Add Column

```go
err := db.AddColumn("Email", "")
```

---

### Remove Column

```go
err := db.RemoveColumn("Email")
```

---

### Sheet Management

```go
sheetNames, err := db.GetAllSheetNames()
exists, err := db.IsSheetExists("Users")
```

---

## 📁 File Structure

- `exceldb.go` – Core logic of the ExcelDB
- `Row` – Type alias for `map[string]interface{}` to represent each data entry

---

## 🧾 Requirements

- Go 1.23.5+
- Excel `.xlsx` files (handled via `github.com/xuri/excelize/v2`)

---

## 📝 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.