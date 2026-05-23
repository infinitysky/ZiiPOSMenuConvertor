package convert

import (
	"fmt"
	"strings"
	"time"

	"menu-converter-v4-go/internal/excelutil"
)

var cnToEn = map[string]string{
	"产品代码":  "ItemCode",
	"名称1":   "Description1",
	"名称2":   "Description2",
	"名称3":   "Description3",
	"名称4":   "Description4",
	"分类":    "Category",
	"菜单组":   "MenuGroup",
	"税率":    "TaxRate",
	"价格1":   "Price",
	"价格2":   "Price2",
	"价格3":   "Price3",
	"价格4":   "Price4",
	"子名称1":  "SubDescription",
	"子名称2":  "SubDescription1",
	"子名称3":  "SubDescription2",
	"子名称4":  "SubDescription3",
	"产品统计组": "ItemGroup",
	"指令":    "Instruction",
}

func NormalizeSource(rows []map[string]any) []map[string]any {
	if len(rows) == 0 {
		return rows
	}
	out := make([]map[string]any, len(rows))
	for i, row := range rows {
		m := make(map[string]any, len(row))
		for k, v := range row {
			if en, ok := cnToEn[k]; ok {
				m[en] = v
			} else {
				m[k] = v
			}
		}
		out[i] = m
	}

	if !excelutil.HasColumn(out, "MenuGroup") {
		for i := range out {
			out[i]["MenuGroup"] = "Default"
		}
	} else {
		for i := range out {
			if excelutil.SafeString(out[i]["MenuGroup"], "") == "" {
				out[i]["MenuGroup"] = "Default"
			}
		}
	}

	if excelutil.HasColumn(out, "ItemCode") {
		norm, instr := 1, 1
		for i := range out {
			if excelutil.SafeString(out[i]["ItemCode"], "") != "" {
				continue
			}
			if excelutil.Bool(out[i]["Instruction"], false) {
				out[i]["ItemCode"] = fmt.Sprintf("I%03d", instr)
				instr++
			} else {
				out[i]["ItemCode"] = fmt.Sprintf("%04d", norm)
				norm++
			}
		}
	}
	return out
}

func ProcessMenuGroup(source []map[string]any, tpl *excelutil.Sheet) (*excelutil.Sheet, map[string]string) {
	base := excelutil.BaseRow(tpl)
	groups := excelutil.UniqueValues(source, "MenuGroup")
	rows := make([]map[string]any, 0, len(groups))
	codeMap := make(map[string]string, len(groups))
	for i, g := range groups {
		row := excelutil.CloneRow(base)
		code := fmt.Sprintf("%02d", i)
		row["Code"] = code
		row["Description"] = g
		row["CultureDescription"] = g
		row["OrderIndex"] = i
		rows = append(rows, row)
		codeMap[g] = code
	}
	return &excelutil.Sheet{Name: tpl.Name, Headers: tpl.Headers, Rows: rows}, codeMap
}

func ProcessCategory(source []map[string]any, tpl *excelutil.Sheet, mgCode map[string]string) *excelutil.Sheet {
	base := excelutil.BaseRow(tpl)
	cats := excelutil.UniqueValues(source, "Category")
	catToMG := make(map[string]string)
	for _, r := range source {
		cat := excelutil.SafeString(r["Category"], "")
		if _, ok := catToMG[cat]; !ok {
			catToMG[cat] = excelutil.SafeString(r["MenuGroup"], "Default")
		}
	}
	rows := make([]map[string]any, 0, len(cats))
	for i, cat := range cats {
		row := excelutil.CloneRow(base)
		mgName := catToMG[cat]
		if mgName == "" {
			mgName = "Default"
		}
		mgCodeVal := mgCode[mgName]
		if mgCodeVal == "" {
			mgCodeVal = "00"
		}
		row["Code"] = fmt.Sprintf("%03d", i+1)
		row["MenuGroupCode"] = mgCodeVal
		row["Category"] = cat
		row["CultureCategory"] = cat
		row["Enable"] = true
		row["OrderIndex"] = i
		row["CategoryGroupSort"] = fmt.Sprintf("%s,%d", mgCodeVal, i)
		row["MenuGroupList"] = mgCodeVal
		rows = append(rows, row)
	}
	return &excelutil.Sheet{Name: tpl.Name, Headers: tpl.Headers, Rows: rows}
}

func ProcessItem(source []map[string]any, tpl *excelutil.Sheet, menuGroupCode string) *excelutil.Sheet {
	base := excelutil.BaseRow(tpl)
	srcCols := make(map[string]struct{})
	for _, c := range excelutil.ColumnNames(source) {
		srcCols[c] = struct{}{}
	}
	catPos := make(map[string]int)
	rows := make([]map[string]any, 0, len(source))

	for i, src := range source {
		row := excelutil.CloneRow(base)
		row["ItemCode"] = excelutil.SafeString(src["ItemCode"], fmt.Sprintf("%04d", i+1))
		desc1 := excelutil.SafeString(src["Description1"], "")
		row["Description1"] = desc1
		row["Description2"] = excelutil.SafeString(src["Description2"], "")
		row["Description3"] = excelutil.SafeString(src["Description3"], "")
		row["Description4"] = excelutil.SafeString(src["Description4"], "")
		row["CultureDescription"] = desc1

		cat := excelutil.SafeString(src["Category"], "")
		row["Category"] = cat
		pos := catPos[cat]
		row["MenuItemCategorySort"] = fmt.Sprintf("%s,%d", cat, pos)
		catPos[cat] = pos + 1

		mainPrice := excelutil.Num(src["Price"], 0)
		row["Price"] = mainPrice
		if _, ok := srcCols["Price1"]; ok && src["Price1"] != nil && fmt.Sprint(src["Price1"]) != "" {
			row["Price1"] = excelutil.Num(src["Price1"], mainPrice)
		} else {
			row["Price1"] = mainPrice
		}
		row["Price2"] = excelutil.Num(src["Price2"], 0)
		row["Price3"] = excelutil.Num(src["Price3"], 0)
		if _, ok := srcCols["Price4"]; ok {
			row["HappyHourPrice4"] = excelutil.Num(src["Price4"], 0)
		}

		row["OnlinePrice1"] = 0
		row["OnlinePrice2"] = 0
		row["OnlinePrice3"] = 0
		row["OnlinePrice4"] = 0

		row["SubDescription"] = excelutil.SafeString(src["SubDescription"], "")
		row["SubDescription1"] = excelutil.SafeString(src["SubDescription1"], "")
		row["SubDescription2"] = excelutil.SafeString(src["SubDescription2"], "")
		row["SubDescription3"] = excelutil.SafeString(src["SubDescription3"], "")

		subFields := []string{"SubDescription", "SubDescription1", "SubDescription2", "SubDescription3"}
		priceFields := []string{"Price1", "Price2", "Price3", "Price4"}
		hasSub := false
		for _, c := range subFields {
			if _, ok := srcCols[c]; ok && excelutil.SafeString(src[c], "") != "" {
				hasSub = true
				break
			}
		}
		hasMulti := false
		for _, c := range priceFields {
			if _, ok := srcCols[c]; ok && excelutil.Num(src[c], 0) > 0 {
				hasMulti = true
				break
			}
		}
		row["Multiple"] = hasSub && hasMulti

		row["TaxRate"] = excelutil.Num(src["TaxRate"], excelutil.Num(base["TaxRate"], 10))
		itemGroup := excelutil.SafeString(src["ItemGroup"], "")
		if itemGroup == "" {
			itemGroup = excelutil.SafeString(base["ItemGroup"], "OTHERS")
		}
		row["ItemGroup"] = itemGroup
		row["Instruction"] = excelutil.Bool(src["Instruction"], false)

		for _, c := range []string{"Scalable", "OpenPrice", "OnlineStatus", "QRCodeStatus"} {
			if _, ok := srcCols[c]; ok && src[c] != nil && fmt.Sprint(src[c]) != "" {
				row[c] = excelutil.Bool(src[c], excelutil.Bool(row[c], false))
			}
		}
		for _, c := range []string{"PrinterPort1", "PrinterPort2", "PrinterPort3", "PrinterPort4", "HappyHourPrice1", "HappyHourPrice2", "HappyHourPrice3"} {
			if _, ok := srcCols[c]; ok && src[c] != nil && fmt.Sprint(src[c]) != "" {
				row[c] = excelutil.Num(src[c], 0)
			}
		}
		row["OrderIndex"] = i
		rows = append(rows, row)
	}
	return &excelutil.Sheet{Name: tpl.Name, Headers: tpl.Headers, Rows: rows}
}

type ExportOptions struct {
	SourceFile    string
	TemplateFile  string
	OutputDir     string
	ShopName      string
	MenuGroupCode string
}

func ProcessMenu(opts ExportOptions) (string, error) {
	if opts.MenuGroupCode == "" {
		opts.MenuGroupCode = "00"
	}
	sheets, order, err := excelutil.ReadWorkbook(opts.SourceFile)
	if err != nil {
		return "", err
	}
	first := order[0]
	source := NormalizeSource(sheets[first].Rows)

	tplSheets, tplOrder, err := excelutil.ReadWorkbook(opts.TemplateFile)
	if err != nil {
		return "", err
	}

	mgSheet, mgMap := ProcessMenuGroup(source, tplSheets["MenuGroupTable"])
	tplSheets["MenuGroupTable"] = mgSheet
	tplSheets["Category"] = ProcessCategory(source, tplSheets["Category"], mgMap)
	tplSheets["MenuItem"] = ProcessItem(source, tplSheets["MenuItem"], opts.MenuGroupCode)

	overview := &excelutil.Sheet{
		Name:    "Overview",
		Headers: []string{"ExportTime"},
		Rows: []map[string]any{
			{"ExportTime": time.Now().Format("2006-01-02 15:04")},
		},
	}

	exportOrder := []string{"Overview"}
	for _, name := range tplOrder {
		if name == "Overview" {
			continue
		}
		exportOrder = append(exportOrder, name)
	}
	allSheets := map[string]*excelutil.Sheet{"Overview": overview}
	for k, v := range tplSheets {
		allSheets[k] = v
	}

	postProcess(allSheets, opts.MenuGroupCode)

	prefix := ""
	if strings.TrimSpace(opts.ShopName) != "" {
		prefix = "-" + strings.TrimSpace(opts.ShopName)
	}
	outFile := fmt.Sprintf("%s\\export_FullMenu%s-%s.xlsx", opts.OutputDir, prefix, time.Now().Format("20060102150405"))
	textCols := map[string][]string{
		"MenuGroupTable": {"A"},
		"Course":         {"A"},
		"Category":       {"A", "B"},
	}
	if err := excelutil.WriteExport(outFile, exportOrder, allSheets, textCols); err != nil {
		return "", err
	}
	return outFile, nil
}

func postProcess(sheets map[string]*excelutil.Sheet, menuGroupCode string) {
	if s, ok := sheets["MenuGroupTable"]; ok {
		for _, row := range s.Rows {
			row["Code"] = excelutil.FormatIntCode(row["Code"], 2)
		}
	}
	if s, ok := sheets["Course"]; ok {
		for i, row := range s.Rows {
			row["CourseCode"] = fmt.Sprintf("%02d", i+1)
		}
	}
	if s, ok := sheets["Category"]; ok {
		for _, row := range s.Rows {
			row["Code"] = excelutil.FormatIntCode(row["Code"], 3)
			if excelutil.SafeString(row["MenuGroupCode"], "") == "" {
				row["MenuGroupCode"] = menuGroupCode
			} else {
				row["MenuGroupCode"] = strings.TrimSpace(fmt.Sprint(row["MenuGroupCode"]))
			}
			if excelutil.SafeString(row["MenuGroupList"], "") == "" {
				row["MenuGroupList"] = menuGroupCode
			} else {
				row["MenuGroupList"] = strings.TrimSpace(fmt.Sprint(row["MenuGroupList"]))
			}
		}
	}
}

func ProcessFromRows(source []map[string]any, templateFile, outputDir string) (string, error) {
	source = NormalizeSource(source)
	tplSheets, tplOrder, err := excelutil.ReadWorkbook(templateFile)
	if err != nil {
		return "", err
	}
	mgSheet, mgMap := ProcessMenuGroup(source, tplSheets["MenuGroupTable"])
	tplSheets["MenuGroupTable"] = mgSheet
	tplSheets["Category"] = ProcessCategory(source, tplSheets["Category"], mgMap)
	tplSheets["MenuItem"] = ProcessItem(source, tplSheets["MenuItem"], "00")

	overview := &excelutil.Sheet{
		Name:    "Overview",
		Headers: []string{"ExportTime"},
		Rows: []map[string]any{
			{"ExportTime": time.Now().Format("2006-01-02 15:04")},
		},
	}
	exportOrder := []string{"Overview"}
	for _, name := range tplOrder {
		if name == "Overview" {
			continue
		}
		exportOrder = append(exportOrder, name)
	}
	allSheets := map[string]*excelutil.Sheet{"Overview": overview}
	for k, v := range tplSheets {
		allSheets[k] = v
	}
	postProcess(allSheets, "00")
	outFile := fmt.Sprintf("%s\\export_FullMenu-%s.xlsx", outputDir, time.Now().Format("20060102150405"))
	textCols := map[string][]string{
		"MenuGroupTable": {"A"},
		"Course":         {"A"},
		"Category":       {"A", "B"},
	}
	if err := excelutil.WriteExport(outFile, exportOrder, allSheets, textCols); err != nil {
		return "", err
	}
	return outFile, nil
}
