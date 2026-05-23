package pe

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Item struct {
	PEID      string  `json:"pe_id"`
	MenuGroup string  `json:"menu_group"`
	Category  string  `json:"category"`
	Name      string  `json:"name"`
	Price     float64 `json:"price"`
	TaxRate   float64 `json:"tax_rate"`
	Status    string  `json:"status"`
	RowIdx    int     `json:"row_idx"`
	HasImage  bool    `json:"has_image"`
}

var (
	priceCleanRe = regexp.MustCompile(`[¥￥\s,，]`)
	taxRateRe    = regexp.MustCompile(`(\d+(?:\.\d+)?)\s*%`)
)

func ParsePrice(s string) float64 {
	if s == "" {
		return 0
	}
	cleaned := priceCleanRe.ReplaceAllString(s, "")
	f, err := strconv.ParseFloat(cleaned, 64)
	if err != nil {
		return 0
	}
	return f
}

func ParseTaxRate(s string) float64 {
	if s == "" {
		return 10
	}
	m := taxRateRe.FindStringSubmatch(s)
	if len(m) > 1 {
		f, _ := strconv.ParseFloat(m[1], 64)
		return f
	}
	return 10
}

func ParseCategoryPath(catPath string) (menuGroup, category string) {
	catPath = strings.TrimSpace(catPath)
	if catPath == "" {
		return "Default", "Default"
	}
	dash := strings.Index(catPath, "-")
	if dash < 0 {
		jp := strings.Split(catPath, "/")[0]
		jp = strings.TrimSpace(jp)
		if jp == "" {
			jp = "Default"
		}
		return jp, jp
	}
	mgPart := strings.TrimSpace(catPath[:dash])
	catPart := strings.TrimSpace(catPath[dash+1:])
	mgJP := strings.TrimSpace(strings.Split(mgPart, "/")[0])
	catJP := strings.TrimSpace(strings.Split(catPart, "/")[0])
	if mgJP == "" {
		mgJP = "Default"
	}
	if catJP == "" {
		catJP = "Default"
	}
	return mgJP, catJP
}

func hasHeaderRow(f *excelize.File, sheet string) bool {
	v, err := f.GetCellValue(sheet, "A1")
	if err != nil || strings.TrimSpace(v) == "" {
		return false
	}
	if _, err := strconv.ParseFloat(strings.TrimSpace(v), 64); err == nil {
		return false
	}
	keywords := map[string]struct{}{
		"id": {}, "code": {}, "itemcode": {}, "item": {}, "name": {}, "category": {},
		"price": {}, "tax": {}, "status": {}, "商品": {}, "名称": {}, "分类": {}, "价格": {},
	}
	_, ok := keywords[strings.ToLower(strings.TrimSpace(v))]
	return ok
}

func ReadMenu(filePath string) ([]Item, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	sheet := f.GetSheetName(f.GetActiveSheetIndex())
	startRow := 1
	if hasHeaderRow(f, sheet) {
		startRow = 2
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}

	imageRows := pictureRows(f, sheet, len(rows))

	var items []Item
	for rowIdx := startRow; rowIdx <= len(rows); rowIdx++ {
		row := rows[rowIdx-1]
		get := func(col int) string {
			if col-1 < len(row) {
				return strings.TrimSpace(row[col-1])
			}
			return ""
		}
		name := get(3)
		if name == "" {
			continue
		}
		mg, cat := ParseCategoryPath(get(2))
		statusRaw := get(8)
		status := "hide"
		switch statusRaw {
		case "展示", "表示", "show", "Show":
			status = "show"
		}
		items = append(items, Item{
			PEID:      get(1),
			MenuGroup: mg,
			Category:  cat,
			Name:      name,
			Price:     ParsePrice(get(5)),
			TaxRate:   ParseTaxRate(get(6)),
			Status:    status,
			RowIdx:    rowIdx,
			HasImage:  imageRows[rowIdx],
		})
	}
	return items, nil
}

func pictureRows(f *excelize.File, sheet string, maxRow int) map[int]bool {
	out := make(map[int]bool)
	for row := 1; row <= maxRow; row++ {
		cell, _ := excelize.CoordinatesToCellName(4, row) // column D
		pics, err := f.GetPictures(sheet, cell)
		if err == nil && len(pics) > 0 {
			out[row] = true
		}
	}
	return out
}

func sanitizeFilename(name string) string {
	name = strings.TrimSpace(name)
	name = strings.ReplaceAll(name, "\u3000", " ")
	replacer := strings.NewReplacer(
		"<", "_", ">", "_", ":", "_", "\"", "_", "/", "_", "\\", "_", "|", "_", "?", "_", "*", "_",
	)
	name = replacer.Replace(name)
	name = strings.Trim(name, ". ")
	if name == "" {
		name = "unnamed"
	}
	if len(name) > 120 {
		name = name[:120]
	}
	return name
}

func ExtractImages(filePath string, items []Item, outputDir string) (int, error) {
	picsDir := filepath.Join(outputDir, "pics")
	if err := os.MkdirAll(picsDir, 0o755); err != nil {
		return 0, err
	}

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return 0, err
	}
	defer f.Close()

	sheet := f.GetSheetName(f.GetActiveSheetIndex())
	rowToName := make(map[int]string)
	for _, item := range items {
		rowToName[item.RowIdx] = item.Name
	}

	saved := 0
	nameCounter := make(map[string]int)
	for _, item := range items {
		cell, _ := excelize.CoordinatesToCellName(4, item.RowIdx)
		pics, err := f.GetPictures(sheet, cell)
		if err != nil || len(pics) == 0 {
			continue
		}
		pic := pics[0]
		row := item.RowIdx
		itemName := rowToName[row]
		if itemName == "" {
			itemName = fmt.Sprintf("row_%d", row)
		}
		safeName := sanitizeFilename(itemName)
		if n, ok := nameCounter[safeName]; ok {
			nameCounter[safeName] = n + 1
			safeName = fmt.Sprintf("%s_%d", safeName, n+1)
		} else {
			nameCounter[safeName] = 0
		}

		ext := strings.ToLower(pic.Extension)
		if ext == "" {
			ext = ".png"
		}
		if !strings.HasPrefix(ext, ".") {
			ext = "." + ext
		}
		if ext == ".jpeg" {
			ext = ".jpg"
		}
		outPath := filepath.Join(picsDir, safeName+ext)
		if err := os.WriteFile(outPath, pic.File, 0o644); err != nil {
			continue
		}
		saved++
	}
	return saved, nil
}

func ItemCode(item Item, idx int) string {
	if strings.TrimSpace(item.PEID) != "" {
		return strings.TrimSpace(item.PEID)
	}
	return fmt.Sprintf("%04d", idx+1)
}

func ToSourceRows(items []Item) []map[string]any {
	rows := make([]map[string]any, len(items))
	for i, item := range items {
		rows[i] = map[string]any{
			"ItemCode":     ItemCode(item, i),
			"Description1": item.Name,
			"Description2": "",
			"Description3": "",
			"Description4": "",
			"Category":     item.Category,
			"MenuGroup":    item.MenuGroup,
			"TaxRate":      item.TaxRate,
			"Price":        item.Price,
			"Price2":       0,
			"Price3":       0,
			"Instruction":  false,
		}
	}
	return rows
}
