package main

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/sqweek/dialog"
)

func checkError(message string, err error) {
	if err != nil {
		log.Fatal(message, err)
	}
}

func main() {

	filename, err := dialog.File().Filter("All Files", "*").Load()

	// fmt.Println(filename)
	// basename := filepath.Base(filename)
	// extension := filepath.Ext(basename)
	// name := filename[0 : len(filename)-len(extension)]

	dir, err := filepath.Abs(filepath.Dir(filename))
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return
	}

	// 디렉터리 가져오기!!
	fmt.Println(dir)

	for index, name := range f.GetSheetMap() {

		// 	// fmt.Println(index, name)
		f.SetActiveSheet(index)

		savepath := filepath.Join(dir, name+".csv")

		file, err := os.Create(savepath)
		checkError("Cannot create file", err)
		// defer file.Close()

		writer := csv.NewWriter(file)
		// temp_rows := [][]string{}

		rows := f.GetRows(name)

		for _, row := range rows {
			for _, colCell := range row {
				writer.Write([]string{colCell})
				writer.Flush()
				// temp_rows = append(temp_r, colCell)
				// fmt.Print(colCell, "\t")
			}
			// fmt.Println()
		}

		// fmt.Println(temp_rows)

		// writer.Flush()

		// f.SaveAs(savepath)
		// 	// fmt.Println(savepath)

	}

}
