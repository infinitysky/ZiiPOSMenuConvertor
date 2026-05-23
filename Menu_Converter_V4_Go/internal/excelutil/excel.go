package excelutil

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Sheet struct {
	Name    string
	Headers []string
	Rows    []map[string]any
}

func ReadWorkbook(path string) (map[string]*Sheet, []string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, nil, err
	}
	defer f.Close()

	order := f.GetSheetList()
	out := make(map[string]*Sheet, len(order))
	for _, name := range order {
		sheet, err := readSheet(f, name)
		if err != nil {
			return nil, nil, err
		}
		out[name] = sheet
	}
	return out, order, nil
}

func readSheet(f *excelize.File, name string) (*Sheet, error) {
	rows, err := f.GetRows(name)
	if err != nil {
		return nil, err
	}
	if len(rows) == 0 {
		return &Sheet{Name: name}, nil
	}
	headers := make([]string, len(rows[0]))
	for i, h := range rows[0] {
		headers[i] = strings.TrimSpace(h)
	}
	data := make([]map[string]any, 0, len(rows)-1)
	for _, row := range rows[1:] {
		m := make(map[string]any, len(headers))
		for i, h := range headers {
			if h == "" {
				continue
			}
			var v any = ""
			if i < len(row) {
				v = row[i]
			}
			m[h] = v
		}
		data = append(data, m)
	}
	return &Sheet{Name: name, Headers: headers, Rows: data}, nil
}

func CloneRow(base map[string]any) map[string]any {
	out := make(map[string]any, len(base))
	for k, v := range base {
		out[k] = v
	}
	return out
}

func BaseRow(sheet *Sheet) map[string]any {
	if sheet == nil || len(sheet.Rows) == 0 {
		return map[string]any{}
	}
	return CloneRow(sheet.Rows[0])
}

func SheetToTable(sheet *Sheet) [][]any {
	if sheet == nil {
		return nil
	}
	table := make([][]any, 0, len(sheet.Rows)+1)
	header := make([]any, len(sheet.Headers))
	for i, h := range sheet.Headers {
		header[i] = h
	}
	table = append(table, header)
	for _, row := range sheet.Rows {
		line := make([]any, len(sheet.Headers))
		for i, h := range sheet.Headers {
			if v, ok := row[h]; ok {
				line[i] = v
			} else {
				line[i] = ""
			}
		}
		table = append(table, line)
	}
	return table
}

func WriteExport(path string, sheetOrder []string, sheets map[string]*Sheet, textCols map[string][]string) error {
	f := excelize.NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	usedDefault := false

	for i, name := range sheetOrder {
		sheet := sheets[name]
		if sheet == nil {
			continue
		}
		target := name
		if i == 0 && !usedDefault {
			_ = f.SetSheetName(defaultSheet, name)
			usedDefault = true
		} else if name != f.GetSheetName(0) {
			_, _ = f.NewSheet(name)
		}
		table := SheetToTable(sheet)
		for r, row := range table {
			cell, _ := excelize.CoordinatesToCellName(1, r+1)
			_ = f.SetSheetRow(target, cell, &row)
		}
		for _, col := range textCols[name] {
			colIdx, _ := excelize.ColumnNameToNumber(col)
			style, _ := f.NewStyle(&excelize.Style{NumFmt: 49})
			lastRow := len(table)
			if lastRow < 1 {
				continue
			}
			start, _ := excelize.CoordinatesToCellName(colIdx, 2)
			end, _ := excelize.CoordinatesToCellName(colIdx, lastRow)
			_ = f.SetCellStyle(target, start, end, style)
		}
	}

	if !usedDefault {
		_ = f.SetSheetName(defaultSheet, sheetOrder[0])
	}

	return f.SaveAs(path)
}

func CellRow(cell string) (int, error) {
	_, row, err := excelize.CellNameToCoordinates(cell)
	return row, err
}

func FormatIntCode(v any, width int) string {
	n := ToInt(v)
	return fmt.Sprintf("%0*d", width, n)
}

func ToInt(v any) int {
	switch x := v.(type) {
	case int:
		return x
	case int64:
		return int(x)
	case float64:
		return int(x)
	case string:
		s := strings.TrimSpace(x)
		if s == "" {
			return 0
		}
		if n, err := strconv.Atoi(s); err == nil {
			return n
		}
		if f, err := strconv.ParseFloat(s, 64); err == nil {
			return int(f)
		}
	}
	return 0
}

func ToFloat(v any) (float64, bool) {
	switch x := v.(type) {
	case float64:
		return x, true
	case int:
		return float64(x), true
	case int64:
		return float64(x), true
	case string:
		s := strings.TrimSpace(x)
		if s == "" {
			return 0, false
		}
		f, err := strconv.ParseFloat(s, 64)
		return f, err == nil
	default:
		return 0, false
	}
}

func SafeString(v any, def string) string {
	if v == nil {
		return def
	}
	s := strings.TrimSpace(fmt.Sprint(v))
	if s == "" {
		return def
	}
	return s
}

func Num(v any, def float64) float64 {
	if f, ok := ToFloat(v); ok {
		return f
	}
	return def
}

func Bool(v any, def bool) bool {
	if v == nil {
		return def
	}
	s := strings.ToLower(strings.TrimSpace(fmt.Sprint(v)))
	switch s {
	case "true", "1", "yes":
		return true
	case "false", "0", "no":
		return false
	default:
		return def
	}
}

func HasColumn(rows []map[string]any, col string) bool {
	for _, r := range rows {
		if _, ok := r[col]; ok {
			return true
		}
	}
	return false
}

func ColumnNames(rows []map[string]any) []string {
	if len(rows) == 0 {
		return nil
	}
	seen := make(map[string]struct{})
	var cols []string
	for _, r := range rows {
		for k := range r {
			if _, ok := seen[k]; !ok {
				seen[k] = struct{}{}
				cols = append(cols, k)
			}
		}
	}
	return cols
}

func UniqueValues(rows []map[string]any, col string) []string {
	seen := make(map[string]struct{})
	var out []string
	for _, r := range rows {
		v := SafeString(r[col], "")
		if _, ok := seen[v]; !ok {
			seen[v] = struct{}{}
			out = append(out, v)
		}
	}
	return out
}
