package poigo

import (
	"fmt"
	"github.com/Centny/jnigo"
	"testing"
)

func TestRun(t *testing.T) {
	NewHSSFWorkbook()
	NewXSSFWorkbook()
	OpenHSSFWorkbook(nil)
	OpenXSSFWorkbook(nil)
	Check(WbSig)
	Init("/Users/cny/Tmp/HH/poi")
	fmt.Println(jnigo.GVM)
	tTestCheck(t)
	tTestFile(t)
	tTestCreate(t)
	tTestRead(t)
}
func tTestCheck(t *testing.T) {
	Check(WbSig + "ss")
	fmt.Println(Check("java.lang.String"))
}
func tTestRead(t *testing.T) {
	input, err := NewFileInputStream("t.xlsx")
	if err != nil {
		t.Error(err.Error())
		return
	}
	defer input.Close()
	wb, err := OpenXSSFWorkbook(input)
	if err != nil {
		t.Error(err.Error())
		return
	}
	sTestRead(wb, t)
	//
	input, err = NewFileInputStream("t.xls")
	if err != nil {
		t.Error(err.Error())
		return
	}
	defer input.Close()
	wb, err = OpenHSSFWorkbook(input)
	if err != nil {
		t.Error(err.Error())
		return
	}
	sTestRead(wb, t)
	//
	input, err = NewFileInputStream("t.xls")
	if err != nil {
		t.Error(err.Error())
		return
	}
	defer input.Close()
	wb, err = OpenHSSFWorkbook(input)
	if err != nil {
		t.Error(err.Error())
		return
	}
	sTestRead2(wb, t)
	//
	input, err = NewFileInputStream("t.xlsx")
	if err != nil {
		t.Error(err.Error())
		return
	}
	defer input.Close()
	wb, err = OpenXSSFWorkbook(input)
	if err != nil {
		t.Error(err.Error())
		return
	}
	sTestRead3(wb, t)
}
func sTestRead(wb *Workbook, t *testing.T) {
	fmt.Println("Read--------->S")
	fmt.Println("sTestRead---->00")
	fmt.Println(wb.NumberOfSheets())
	fmt.Println(wb.SheetName(1000))
	sheet, err := wb.SheetAt(0)
	if err != nil {
		t.Error(err.Error())
		return
	}
	fmt.Println(wb.SheetAt(100))
	fmt.Println(sheet.CellAt(0, 0))
	fmt.Println(sheet.CellAt(100, 10))
	wb.SetSheetName(0, "name")
	fmt.Println(sheet.SheetName())
	fmt.Println(wb.SheetName(0))
	fmt.Println(wb.Sheet("name"))
	fmt.Println(wb.SheetIndex("name"))
	fmt.Println("sTestRead---->00-1")
	row, err := sheet.RowAt(0)
	if err != nil {
		t.Error(err.Error())
		return
	}
	fmt.Println(row.FirstCellNum())
	fmt.Println(row.LastCellNum())
	fmt.Println(row.PhysicalNumberOfCells())
	fmt.Println(row.RowNum())
	fmt.Println(row.Sheet())
	fmt.Println("sTestRead---->01-1")
	fmt.Println("sTestRead---->01")
	//
	cell, err := row.CellAt(0)
	if err != nil {
		t.Error(err.Error())
		return
	}
	fmt.Println(row.CellAt(10000))
	fmt.Println(cell.String())
	fmt.Println(cell.Row())
	fmt.Println(cell.Sheet())
	fmt.Println(cell.ColumnIndex())
	fmt.Println(cell.RowIndex())
	fmt.Println("sTestRead---->02")
	//
	row, err = sheet.RowAt(1)
	if err != nil {
		t.Error(err.Error())
		return
	}
	row.SetRowNum(3)
	cell, err = row.CellAt(0)
	if err != nil {
		t.Error(err.Error())
		return
	}
	cell.SetCellType(CELL_TYPE_STRING)
	fmt.Println(cell.CellType())
	row.RemoveCell(cell)
	fmt.Println("sTestRead---->03")
	//
	sheet.RemoveRow(row)
	fmt.Println(sheet.Workbook())
	fmt.Println(sheet.PhysicalNumberOfRows())
	fmt.Println(sheet.FirstRowNum())
	fmt.Println(sheet.LastRowNum())
	// fmt.Println(sheet.SheetName())
	//
	fmt.Println("sTestRead---->04")
	// fmt.Println(wb.NumberOfSheets())
	// fmt.Println(wb.SheetName(0))
	// wb.SetSheetName(0, "jjjjjj")
	// fmt.Println(wb.Sheet("jjjjjj"))
	fmt.Println("sTestRead---->05")
	fmt.Println("Read--------->")
}
func sTestRead2(wb *Workbook, t *testing.T) {
	sheet, err := wb.SheetAt(0)
	if err != nil {
		t.Error(err.Error())
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
		t.Error(err.Error())
		return
	}
}
func sTestRead3(wb *Workbook, t *testing.T) {
	sheet, err := wb.SheetAt(0)
	if err != nil {
		t.Error(err.Error())
		return
	}
	cell, err := sheet.CellAt(3, 2)
	if err != nil {
		t.Error(err.Error())
		return
	}
	fmt.Println(cell.Formula())
}
func tTestCreate(t *testing.T) {
	wb, err := NewXSSFWorkbook()
	if err != nil {
		t.Error(err.Error())
		return
	}
	sTestCreate(wb, t)
	//
	wb, err = NewHSSFWorkbook()
	if err != nil {
		t.Error(err.Error())
		return
	}
	sTestCreate(wb, t)
	//
	// wb, err = NewHSSFWorkbook()
	// if err != nil {
	// 	t.Error(err.Error())
	// 	return
	// }
	// sTestCreate2(wb, t)
}
func sTestCreate(wb *Workbook, t *testing.T) {
	fmt.Println("TestCreate--------->S")
	sheet, err := wb.CreateSheet()
	if err != nil {
		t.Error(err.Error())
		return
	}
	wb.RemoveSheetAt(0)
	//
	sheet, err = wb.CreateSheet("abc")
	if err != nil {
		t.Error(err.Error())
		return
	}
	row, err := sheet.CreateRow(0)
	if err != nil {
		t.Error(err.Error())
		return
	}
	//
	cell, err := row.CreateCell(0)
	if err != nil {
		t.Error(err.Error())
		return
	}
	err = cell.SetCellValue("abbb1")
	if err != nil {
		t.Error(err.Error())
		return
	}
	//
	cell, err = row.CreateCell(1)
	if err != nil {
		t.Error(err.Error())
		return
	}
	err = cell.SetCellValue("abbb1")
	if err != nil {
		t.Error(err.Error())
		return
	}
	//
	cell, err = row.CreateCell(2)
	if err != nil {
		t.Error(err.Error())
		return
	}
	date, err := jnigo.GVM.New("java.util.Date", int64(100000))
	if err != nil {
		t.Error(err.Error())
		return
	}
	err = cell.SetCellValue(date)
	if err != nil {
		t.Error(err.Error())
		return
	}
	fmt.Println(cell.Date())
	//
	cell, err = row.CreateCell(3)
	if err != nil {
		t.Error(err.Error())
		return
	}
	err = cell.SetCellValue(false)
	if err != nil {
		t.Error(err.Error())
		return
	}
	fmt.Println(cell.Boolean())
	fmt.Println(cell.String())
	fmt.Println(cell.Numeric())
	fmt.Println(cell.Date())
	//
	cell, err = row.CreateCell(4)
	if err != nil {
		t.Error(err.Error())
		return
	}
	err = cell.SetCellValue(float64(1111))
	if err != nil {
		t.Error(err.Error())
		return
	}
	fmt.Println(cell.Numeric())
	fmt.Println(cell.String())
	fmt.Println(cell.Boolean())
	fmt.Println(cell.Date())
	//
	// cell, err = row.CreateCell(5, CELL_TYPE_FORMULA)
	// if err != nil {
	// 	t.Error(err.Error())
	// 	return
	// }
	// err = cell.SetCellFormula("=A0")
	// if err != nil {
	// 	t.Error(err.Error())
	// 	return
	// }
	// fmt.Println(cell.Formula())
	// fmt.Println("------>")
	//
	out, err := NewFileOutputStream("/tmp/t.tmp")
	if err != nil {
		t.Error(err.Error())
		return
	}
	err = wb.Write(out)
	if err != nil {
		t.Error(err.Error())
		return
	}
	err = out.Close()
	if err != nil {
		t.Error(err.Error())
		return
	}
	fmt.Println("TestCreate--------->")
}

// func sTestCreate2(wb *Workbook, t *testing.T) {
// 	sheet, err := wb.CreateSheet("abc")
// 	if err != nil {
// 		t.Error(err.Error())
// 		return
// 	}
// 	row, err := sheet.CreateRow(0)
// 	if err != nil {
// 		t.Error(err.Error())
// 		return
// 	}
// 	cell, err := row.CreateCell(0, CELL_TYPE_NUMERIC)
// 	if err != nil {
// 		t.Error(err.Error())
// 		return
// 	}
// 	cell.SetCellValue("1")
// 	cell, err = row.CreateCell(1, CELL_TYPE_NUMERIC)
// 	if err != nil {
// 		t.Error(err.Error())
// 		return
// 	}
// 	cell.SetCellValue("1")
// 	cell, err = row.CreateCell(2, CELL_TYPE_FORMULA)
// 	if err != nil {
// 		t.Error(err.Error())
// 		return
// 	}
// 	fmt.Println(cell.SetCellFormula("=A1"))
// 	ev, err := cell.Evaluate()
// 	if err != nil {
// 		t.Error(err.Error())
// 		return
// 	}
// 	fmt.Println(ev.String())
// }
