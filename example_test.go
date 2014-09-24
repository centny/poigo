package poigo

import (
	"fmt"
)

func ExamplePoi() {
	Init("<class path for poi>")
	input, err := NewFileInputStream("t.xlsx")
	if err != nil {
		return
	}
	defer input.Close()
	wb, err := OpenXSSFWorkbook(input)
	if err != nil {
		return
	}
	sheet, err := wb.SheetAt(0)
	if err != nil {
		return
	}
	err = sheet.Loop(func(r *Row) {
		fmt.Println("xxxxxx000000xxxxx")
		fmt.Println(r)
		fmt.Println(r.PhysicalNumberOfCells())
		r.Loop(func(c *Cell) {
			fmt.Println(c.String())
		})
	})
	if err != nil {
		return
	}
}
