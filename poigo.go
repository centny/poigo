package poigo

import (
	"fmt"
	"github.com/Centny/gwf/util"
	"github.com/Centny/jnigo"
	"time"
)

const (
	CELL_TYPE_NUMERIC = 0
	CELL_TYPE_STRING  = 1
	CELL_TYPE_FORMULA = 2
	CELL_TYPE_BLANK   = 3
	CELL_TYPE_BOOLEAN = 4
	CELL_TYPE_ERROR   = 5
)

const WbSig = "org.apache.poi.ss.usermodel.Workbook"

func Init(path ...string) int {
	return jnigo.Init(jnigo.NewClassPathOption(path...))
}
func chk_vm() error {
	if jnigo.GVM == nil {
		return jnigo.Err("JVM is nil (initial first)")
	} else {
		return nil
	}
}
func Check(sig string) error {
	if err := chk_vm(); err != nil {
		return err
	}
	cls := jnigo.GVM.FindClass(sig)
	if cls == nil {
		return jnigo.Err("class not found(POI library is missing?)")
	} else {
		return nil
	}
}

////////////////////////////////////////////
type Stream interface {
	S() *jnigo.Object
}

////////////////////////////////////////////
func NewXSSFWorkbook() (*Workbook, error) {
	if err := chk_vm(); err != nil {
		return nil, err
	}
	bk, err := jnigo.GVM.NewAs("org.apache.poi.xssf.usermodel.XSSFWorkbook", WbSig)
	return &Workbook{
		Wb: bk,
	}, err
}
func OpenXSSFWorkbook(input Stream) (*Workbook, error) {
	if err := chk_vm(); err != nil {
		return nil, err
	}
	bk, err := jnigo.GVM.NewAs("org.apache.poi.xssf.usermodel.XSSFWorkbook", WbSig, input.S())
	return &Workbook{
		Wb: bk,
	}, err
}
func NewHSSFWorkbook() (*Workbook, error) {
	if err := chk_vm(); err != nil {
		return nil, err
	}
	bk, err := jnigo.GVM.NewAs("org.apache.poi.hssf.usermodel.HSSFWorkbook", WbSig)
	return &Workbook{
		Wb: bk,
	}, err
}
func OpenHSSFWorkbook(input Stream) (*Workbook, error) {
	if err := chk_vm(); err != nil {
		return nil, err
	}
	bk, err := jnigo.GVM.NewAs("org.apache.poi.hssf.usermodel.HSSFWorkbook", WbSig, input.S())
	return &Workbook{
		Wb: bk,
	}, err
}

////////////////////////////////////////////
type Workbook struct {
	Wb      *jnigo.Object
	Formula *FormulaEvaluator
}

func (w *Workbook) SheetAt(idx int) (*Sheet, error) {
	sb, _ := w.Wb.CallObject("getSheetAt", "org.apache.poi.ss.usermodel.Sheet", idx)
	if sb == nil {
		return nil, jnigo.Err("Sheet not found by index(%d)", idx)
	} else {
		return &Sheet{
			W:  w,
			Sb: sb,
		}, nil
	}
}

// func (w *Workbook) ActiveSheetIndex() (int, error) {
// 	return w.Wb.CallInt("getActiveSheetIndex")
// }
// func (w *Workbook) SetActiveSheet(sheetIndex int) error {
// 	return w.Wb.CallVoid("setActiveSheet", sheetIndex)
// }
// func (w *Workbook) FirstVisibleTab() (int, error) {
// 	return w.Wb.CallInt("getFirstVisibleTab")
// }
// func (w *Workbook) SetFirstVisibleTab(sheetIndex int) error {
// 	return w.Wb.CallVoid("setFirstVisibleTab", sheetIndex)
// }
// func (w *Workbook) SetSheetOrder(sheetname string, pos int) error {
// 	return w.Wb.CallVoid("setSheetOrder", sheetname, pos)
// }
// func (w *Workbook) SetSelectedTab(index int) error {
// 	return w.Wb.CallVoid("setSelectedTab", index)
// }
func (w *Workbook) SetSheetName(sheet int, name string) error {
	return w.Wb.CallVoid("setSheetName", sheet, name)
}
func (w *Workbook) SheetName(sheet int) (string, error) {
	sval, err := w.Wb.CallObject("getSheetName", "java.lang.String", sheet)
	if err != nil {
		return "", err
	}
	return sval.AsString(), nil
}
func (w *Workbook) SheetIndex(v interface{}) (int, error) {
	return w.Wb.CallInt("getSheetIndex", v)
}
func (w *Workbook) CreateSheet(v ...interface{}) (*Sheet, error) {
	sb, err := w.Wb.CallObject("createSheet", "org.apache.poi.ss.usermodel.Sheet", v...)
	return &Sheet{
		W:  w,
		Sb: sb,
	}, err
}

// func (w *Workbook) CloneSheet(sheetNum int) (*Sheet, error) {
// 	sb, err := w.Wb.CallObject("CloneSheet", "org.apache.poi.ss.usermodel.Sheet", sheetNum)
// 	return &Sheet{
// 		Sb: sb,
// 	}, err
// }
func (w *Workbook) NumberOfSheets() (int, error) {
	return w.Wb.CallInt("getNumberOfSheets")
}
func (w *Workbook) Sheet(name string) (*Sheet, error) {
	sb, err := w.Wb.CallObject("getSheet", "org.apache.poi.ss.usermodel.Sheet", name)
	return &Sheet{
		W:  w,
		Sb: sb,
	}, err
}
func (w *Workbook) RemoveSheetAt(index int) error {
	return w.Wb.CallVoid("removeSheetAt", index)
}
func (w *Workbook) Write(output Stream) error {
	return w.Wb.CallVoid("write", output.S())
}

// func (w *Workbook) IsHidden() (bool, error) {
// 	return w.Wb.CallBoolean("isHidden")
// }
// func (w *Workbook) SetHidden(hiddenFlag bool) error {
// 	return w.Wb.CallVoid("setHidden", hiddenFlag)
// }
// func (w *Workbook) IsSheetHidden(sheetIx int) (bool, error) {
// 	return w.Wb.CallBoolean("isSheetHidden")
// }
// func (w *Workbook) IsSheetVeryHidden(sheetIx int) (bool, error) {
// 	return w.Wb.CallBoolean("isSheetVeryHidden")
// }
// func (w *Workbook) SetSheetHidden(sheetIx int, hidden interface{}) error {
// 	return w.Wb.CallVoid("setSheetHidden", sheetIx, hidden)
// }
func (w *Workbook) FormulaEvaluator() *FormulaEvaluator {
	if w.Formula == nil {
		w.Formula, _ = w.CreateFormulaEvaluator()
	}
	if w.Formula == nil {
		panic("FormulaEvaluator is nil")
	}
	return w.Formula
}
func (w *Workbook) CreateFormulaEvaluator() (*FormulaEvaluator, error) {
	ch, err := w.Wb.CallObject("getCreationHelper", "org.apache.poi.ss.usermodel.CreationHelper")
	if err != nil {
		return nil, err
	}
	fe, err := ch.CallObject("createFormulaEvaluator", "org.apache.poi.ss.usermodel.FormulaEvaluator")
	return &FormulaEvaluator{
		W:       w,
		Formula: fe,
	}, err
}

////////////////////////////////////////////
type FormulaEvaluator struct {
	W       *Workbook
	Formula *jnigo.Object
}

func (f *FormulaEvaluator) ClearAllCachedResultValues() error {
	return f.Formula.CallVoid("clearAllCachedResultValues")
}
func (f *FormulaEvaluator) NotifySetFormula(cell *Cell) error {
	return f.Formula.CallVoid("notifySetFormula", cell.Cb)
}
func (f *FormulaEvaluator) NotifyDeleteCell(cell *Cell) error {
	return f.Formula.CallVoid("notifyDeleteCell", cell.Cb)
}
func (f *FormulaEvaluator) NotifyUpdateCell(cell *Cell) error {
	return f.Formula.CallVoid("notifyDeleteCell", cell.Cb)
}
func (f *FormulaEvaluator) EvaluateAll() error {
	return f.Formula.CallVoid("evaluateAll")
}
func (f *FormulaEvaluator) Evaluate(cell *Cell) (*CellValue, error) {
	cb, err := f.Formula.CallObject("evaluate", "org.apache.poi.ss.usermodel.CellValue", cell.Cb)
	if cb == nil {
		return nil, jnigo.Err("nil error")
	} else {
		return &CellValue{
			Formula: f,
			C:       cell,
			Cb:      cb,
		}, err
	}
}
func (f *FormulaEvaluator) EvaluateFormulaCell(cell *Cell) (int, error) {
	return f.Formula.CallInt("evaluateFormulaCell", cell.Cb)
}
func (f *FormulaEvaluator) EvaluateInCell(cell *Cell) (*Cell, error) {
	cb, err := f.Formula.CallObject("evaluateInCell", "org.apache.poi.ss.usermodel.Cell", cell.Cb)
	if cb == nil {
		return nil, jnigo.Err("nil error")
	} else {
		return &Cell{
			R:  cell.R,
			Cb: cb,
		}, err
	}
}
func (f *FormulaEvaluator) SetDebugEvaluationOutputForNextEval(value bool) error {
	return f.Formula.CallVoid("setDebugEvaluationOutputForNextEval", value)
}

////////////////////////////////////////////
type Sheet struct {
	W  *Workbook
	Sb *jnigo.Object
}

func (s *Sheet) RowAt(idx int) (*Row, error) {
	rb, _ := s.Sb.CallObject("getRow", "org.apache.poi.ss.usermodel.Row", idx)
	if rb == nil {
		return nil, jnigo.Err("Row not found by index(%v)", idx)
	} else {
		return &Row{
			S:  s,
			Rb: rb,
		}, nil
	}
}
func (s *Sheet) CellAt(r, c int) (*Cell, error) {
	row, err := s.RowAt(r)
	fmt.Println(row)
	if err != nil {
		return nil, err
	}
	return row.CellAt(c)
}
func (s *Sheet) CellRef(ref string) (*Cell, error) {
	cr, err := jnigo.GVM.New("org.apache.poi.ss.util.CellReference", ref)
	if err != nil {
		return nil, err
	}
	r, _ := cr.CallInt("getRow")
	c, _ := cr.CallInt("getCol")
	return s.CellAt(r, c)
}
func (s *Sheet) CreateRow(rownum int) (*Row, error) {
	rb, err := s.Sb.CallObject("createRow", "org.apache.poi.ss.usermodel.Row", rownum)
	return &Row{
		S:  s,
		Rb: rb,
	}, err
}
func (s *Sheet) RemoveRow(row *Row) error {
	return s.Sb.CallVoid("createRow", row)
}
func (s *Sheet) PhysicalNumberOfRows() (int, error) {
	return s.Sb.CallInt("getPhysicalNumberOfRows")
}
func (s *Sheet) FirstRowNum() (int, error) {
	return s.Sb.CallInt("getFirstRowNum")
}
func (s *Sheet) LastRowNum() (int, error) {
	return s.Sb.CallInt("getLastRowNum")
}
func (s *Sheet) Workbook() (*Workbook, error) {
	wb, err := s.Sb.CallObject("getWorkbook", "org.apache.poi.ss.usermodel.Workbook")
	return &Workbook{
		Wb: wb,
	}, err
}
func (s *Sheet) SheetName() string {
	sval, _ := s.Sb.CallObject("getSheetName", "java.lang.String")
	return sval.ToString()
}
func (s *Sheet) Loop(f func(r *Row)) error {
	it, err := s.Sb.CallObject("rowIterator", "java.util.Iterator")
	if err != nil {
		return err
	}
	for {
		bv, _ := it.CallBoolean("hasNext")
		if !bv {
			break
		}
		ov, _ := it.CallObject("next", "java.lang.Object")
		ov, _ = ov.As("org.apache.poi.ss.usermodel.Row")
		f(&Row{
			S:  s,
			Rb: ov,
		})
	}
	return nil
}

////////////////////////////////////////////
type Row struct {
	S  *Sheet
	Rb *jnigo.Object
}

func (r *Row) Sheet() (*Sheet, error) {
	sb, err := r.Rb.CallObject("getSheet", "org.apache.poi.ss.usermodel.Sheet")
	return &Sheet{
		W:  r.S.W,
		Sb: sb,
	}, err
}
func (r *Row) CellAt(column int) (*Cell, error) {
	cb, _ := r.Rb.CallObject("getCell", "org.apache.poi.ss.usermodel.Cell", column)
	if cb == nil {
		return nil, jnigo.Err("Cell not found by index(%v)", column)
	} else {
		return &Cell{
			R:  r,
			Cb: cb,
		}, nil
	}
}
func (r *Row) CreateCell(col_type ...interface{}) (*Cell, error) {
	cb, err := r.Rb.CallObject("createCell", "org.apache.poi.ss.usermodel.Cell", col_type...)
	return &Cell{
		R:  r,
		Cb: cb,
	}, err
}
func (r *Row) RemoveCell(cell *Cell) error {
	return r.Rb.CallVoid("removeCell", cell)
}
func (r *Row) SetRowNum(num int) error {
	return r.Rb.CallVoid("setRowNum", num)
}
func (r *Row) RowNum() (int, error) {
	return r.Rb.CallInt("getRowNum")
}
func (r *Row) FirstCellNum() (int16, error) {
	return r.Rb.CallShort("getFirstCellNum")
}
func (r *Row) LastCellNum() (int16, error) {
	return r.Rb.CallShort("getLastCellNum")
}
func (r *Row) PhysicalNumberOfCells() (int, error) {
	return r.Rb.CallInt("getPhysicalNumberOfCells")
}

func (r *Row) Loop(f func(c *Cell)) error {
	it, err := r.Rb.CallObject("cellIterator", "java.util.Iterator")
	if err != nil {
		return err
	}
	for {
		bv, _ := it.CallBoolean("hasNext")
		if !bv {
			break
		}
		ov, _ := it.CallObject("next", "java.lang.Object")
		ov, _ = ov.As("org.apache.poi.ss.usermodel.Cell")
		f(&Cell{
			R:  r,
			Cb: ov,
		})
	}
	return nil
}

////////////////////////////////////////////
type Cell struct {
	R  *Row
	Cb *jnigo.Object
}

func (c *Cell) ColumnIndex() (int, error) {
	return c.Cb.CallInt("getColumnIndex")
}
func (c *Cell) RowIndex() (int, error) {
	return c.Cb.CallInt("getRowIndex")
}
func (c *Cell) Sheet() (*Sheet, error) {
	sb, err := c.Cb.CallObject("getSheet", "org.apache.poi.ss.usermodel.Sheet")
	return &Sheet{
		W:  c.R.S.W,
		Sb: sb,
	}, err
}
func (c *Cell) Row() (*Row, error) {
	rb, err := c.Cb.CallObject("getRow", "org.apache.poi.ss.usermodel.Row")
	return &Row{
		S:  c.R.S,
		Rb: rb,
	}, err
}
func (c *Cell) CellType() (int, error) {
	return c.Cb.CallInt("getCellType")
}
func (c *Cell) SetCellType(typ int) error {
	return c.Cb.CallVoid("SetCellType", typ)
}
func (c *Cell) SetCellValue(val interface{}) error {
	return c.Cb.CallVoid("setCellValue", val)
}
func (c *Cell) SetCellFormula(formula string) error {
	return c.Cb.CallVoid("setCellFormula", formula)
}

// func (c *Cell) CellFormula() (string, error) {
// 	sval, err := c.Cb.CallObject("getCellFormula", "java.lang.String")
// 	if err != nil {
// 		return "", err
// 	}
// 	return sval.AsString(), nil
// }
func (c *Cell) String() (string, error) {
	sval, err := c.Cb.CallObject("getStringCellValue", "java.lang.String")
	if err != nil {
		return "", err
	}
	return sval.AsString(), nil
}
func (c *Cell) Numeric() (float64, error) {
	return c.Cb.CallDouble("getNumericCellValue")
}
func (c *Cell) Boolean() (bool, error) {
	return c.Cb.CallBoolean("getBooleanCellValue")
}
func (c *Cell) Date() (time.Time, error) {
	date, err := c.Cb.CallObject("getDateCellValue", "java.util.Date")
	if err != nil {
		return time.Now(), err
	}
	ts, err := date.CallLong("getTime")
	return util.Time(ts), err
}
func (c *Cell) Formula() (string, error) {
	cv, _ := c.Evaluate()
	return cv.FormatAsString()
}

//
func (c *Cell) NotifySetFormula() error {
	return c.R.S.W.FormulaEvaluator().NotifySetFormula(c)
}
func (c *Cell) NotifyDeleteCell() error {
	return c.R.S.W.FormulaEvaluator().NotifyDeleteCell(c)
}
func (c *Cell) NotifyUpdateCell() error {
	return c.R.S.W.FormulaEvaluator().NotifyUpdateCell(c)
}
func (c *Cell) EvaluateAll() error {
	return c.R.S.W.FormulaEvaluator().EvaluateAll()
}
func (c *Cell) Evaluate() (*CellValue, error) {
	return c.R.S.W.FormulaEvaluator().Evaluate(c)
}
func (c *Cell) EvaluateFormulaCell() (int, error) {
	return c.R.S.W.FormulaEvaluator().EvaluateFormulaCell(c)
}
func (c *Cell) EvaluateInCell() (*Cell, error) {
	return c.R.S.W.FormulaEvaluator().EvaluateInCell(c)
}

////////////////////////////////////////////

type CellValue struct {
	Formula *FormulaEvaluator
	C       *Cell
	Cb      *jnigo.Object
}

func (c *CellValue) Boolean() (bool, error) {
	return c.Cb.CallBoolean("getBooleanValue")
}
func (c *CellValue) Number() (float64, error) {
	return c.Cb.CallDouble("getNumberValue")
}
func (c *CellValue) String() (string, error) {
	return c.Cb.CallString("getStringValue")
}
func (c *CellValue) FormatAsString() (string, error) {
	return c.Cb.CallString("formatAsString")
}
