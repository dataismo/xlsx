package xlsx

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"reflect"
	"strconv"
	"time"
	"unicode/utf8"
)

type Sheet struct {
	name       string
	indexRow   int
	lastCol    int
	xlsx       *Xlsx
	startIndex int
}

type styles struct {
	StoneBg     int
	MoneyFormat int
	Bold        int
}

type Xlsx struct {
	file         *excelize.File
	Sheets       map[string]*Sheet
	defaultSheet string
	Styles       *styles
}

// New crea el wapper de excel, renombra la hoja por defecto
func New(sheetname string, startindex int) *Xlsx {

	excel := &Xlsx{
		file:         excelize.NewFile(),
		Sheets:       make(map[string]*Sheet),
		defaultSheet: sheetname,
		Styles:       &styles{},
	}

	f := excel.file

	excel.Styles.StoneBg, _ = f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{"#F4F4F5"},
		},
	})

	excel.Styles.MoneyFormat, _ = f.NewStyle(&excelize.Style{
		NumFmt: 186,
	})

	excel.Styles.Bold, _ = f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
	})

	defaultSheet := excel.file.GetSheetName(0)
	if err := excel.file.SetSheetName(defaultSheet, sheetname); err != nil {
		panic(err)
	}

	sheet := &Sheet{name: sheetname, indexRow: startindex, xlsx: excel, startIndex: startindex}
	excel.Sheets[sheetname] = sheet

	return excel
}

func (x *Xlsx) GetDefaultSheet() *Sheet {
	return x.Sheets[x.defaultSheet]
}

// SaveFile Guarda el archivo segun el path dado
func (x *Xlsx) SaveFile(path string) error {
	return x.file.SaveAs(path)
}

func (x *Xlsx) GetExcelize() *excelize.File {
	return x.file
}

// NewSheet crea una nueva hoja para el excel
func (x *Xlsx) NewSheet(name string, indexRow int) (*Sheet, error) {
	_, err := x.file.NewSheet(name)
	if err != nil {
		return nil, err
	}
	x.Sheets[name] = &Sheet{name: name, xlsx: x, indexRow: indexRow}
	return x.Sheets[name], nil
}

func (sheet *Sheet) AddRowHeader(cols ...interface{}) error {
	if err := sheet.AddRow(cols...); err != nil {
		return err
	}
	if err := sheet.SetRowStyle(sheet.indexRow - 1); err != nil {
		return err
	}
	colstart, colend := sheet.GetCellsFromRow(sheet.indexRow - 1)

	err := sheet.xlsx.file.AutoFilter(sheet.name,
		fmt.Sprintf("%s:%s", colstart, colend),
		nil,
	)

	if err != nil {
		return err
	}

	return nil

}

func (sheet *Sheet) SetRowStyle(rowindex int) error {
	f := sheet.xlsx.file
	colstart, colend := sheet.GetCellsFromRow(rowindex)
	return f.SetCellStyle(sheet.name, colstart, colend, sheet.xlsx.Styles.StoneBg)
}

// GetCellsFromRow regresa la ultima y primera columna segun el index de la fila dado
func (sheet *Sheet) GetCellsFromRow(index int) (string, string) {
	lastcol, _ := excelize.ColumnNumberToName(sheet.lastCol)
	colstart := fmt.Sprintf("A%d", index)
	colend := fmt.Sprintf("%s%d", lastcol, index)
	return colstart, colend
}

// GetCellsFromCol regresa la celda de principio y final de una columna dada
func (sheet *Sheet) GetCellsFromCol(col string) (string, string) {
	cellStart := fmt.Sprintf("%s%d", col, sheet.startIndex+1) // contando la cabecera
	cellEnd := fmt.Sprintf("%s%d", col, sheet.indexRow-1)     // descontando la ultima iteracion, donde se suma al final
	return cellStart, cellEnd
}

// AddRow agrega nueva fila y sus valores, conforme se les pase
func (sheet *Sheet) AddRow(cols ...interface{}) error {
	sheet.lastCol = len(cols)
	for i, val := range cols {
		lastcol := i + 1
		colname, err := excelize.ColumnNumberToName(lastcol)
		if err != nil {
			return err
		}
		rowindex := strconv.Itoa(sheet.indexRow)
		axis := colname + rowindex
		if err = sheet.SetCellValue(axis, val); err != nil {
			return err
		}

	}
	// a final sumamos una nueva fila
	sheet.indexRow++
	return nil
}

// / Subtotal coloca en una celda, la subtotal (suma) de una columuna dada
func (sheet *Sheet) Subtotal(cell string, column string) error {
	f := sheet.xlsx.file
	startCell := fmt.Sprintf("%s%d", column, sheet.startIndex+1) // el mas 1 es por la cabecera
	endCell := fmt.Sprintf("%s%d", column, sheet.indexRow-1)     // por que quedamos un iteracion mas
	formula := fmt.Sprintf("SUBTOTAL(9,%s:%s)", startCell, endCell)
	return f.SetCellFormula(sheet.name, cell, formula)
}

func (sheet *Sheet) SetColumnStyle(styleid int, col string) error {
	cellStart, cellEnd := sheet.GetCellsFromCol(col)
	return sheet.xlsx.file.SetCellStyle(sheet.name, cellStart, cellEnd, styleid)
}

func (sheet *Sheet) SetCellValue(axis string, value any) error {
	var err error
	rType := reflect.TypeOf(value)
	kind := rType.Kind()
	colname := string(axis[0])

	f := sheet.xlsx.file

	sheetname := sheet.name
	val := reflect.ValueOf(value)
	switch kind {
	case reflect.String:
		// en caso de que sea un string contamos el tamaño de la columna para hacer ajustado
		str := val.String()
		length := float64(utf8.RuneCountInString(str) + 4)
		width, err := f.GetColWidth(sheetname, colname)
		if err != nil {
			return err
		}
		if length > width && length < 255 {
			if err = f.SetColWidth(sheetname, colname, colname, length); err != nil {
				return err
			}
		}
		err = f.SetCellStr(sheetname, axis, str)
	case reflect.Float64:
		_ = f.SetColWidth(sheetname, colname, colname, 15)
		err = f.SetCellFloat(sheetname, axis, val.Float(), 2, 64)
	case reflect.Int:
		err = f.SetCellInt(sheetname, axis, int(val.Int()))
	case reflect.Bool:
		res := val.Bool()
		if res {
			err = f.SetCellStr(sheetname, axis, "Si")
		} else {
			err = f.SetCellStr(sheetname, axis, "No")
		}
	}
	if err != nil {
		return err
	}

	return nil
}

func (sheet *Sheet) GetIndexRow() int {
	return sheet.indexRow
}

// Freeze cogela filas y columnas, partiendo desde x y y
func (sheet *Sheet) Freeze(x int, y int) error {
	col, _ := excelize.ColumnNumberToName(y + 1)
	topCell := fmt.Sprintf("%s%d", col, x+1)
	return sheet.xlsx.file.SetPanes(
		sheet.name,
		&excelize.Panes{
			Freeze:      true,
			Split:       true,
			XSplit:      y,
			YSplit:      x,
			ActivePane:  "topRight",
			TopLeftCell: topCell,
		},
	)
}

// Output regresa los bytes del archivo
func (xlsx *Xlsx) Output() ([]byte, error) {
	path := os.TempDir() + "/" + strconv.Itoa(int(time.Now().Unix())) + ".xlsx"
	if err := xlsx.file.SaveAs(path); err != nil {
		return nil, err
	}
	data, err := os.ReadFile(path)
	// boramos el temporal
	if err = os.Remove(path); err != nil {
		return nil, err
	}
	return data, err
}

// SaveTmp retorna el path del del archivo guardado temporalmente
func (xlsx *Xlsx) SaveTmp() (string, error) {
	path := os.TempDir() + "/" + strconv.Itoa(int(time.Now().Unix())) + ".xlsx"
	if err := xlsx.file.SaveAs(path); err != nil {
		return "", err
	}
	return path, nil
}

// SetCellStyle coloca un estylo a una celda
func (sheet *Sheet) SetCellStyle(axis string, styleid int) error {
	return sheet.xlsx.file.SetCellStyle(sheet.name, axis, axis, styleid)
}
