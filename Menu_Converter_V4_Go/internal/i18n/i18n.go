package i18n

var strings = map[string]map[string]string{
	"en": {
		"title":             "ZiiPOS Menu Converter V4.0 (Go)",
		"tab_excel":         "Excel Import",
		"tab_pe":            "PE Menu Import",
		"lbl_menu_file":     "Menu File",
		"lbl_pe_file":       "PE Menu File",
		"lbl_output_folder": "Output Folder",
		"btn_select":        "Select",
		"btn_browse":        "Browse",
		"btn_convert":       "Convert",
		"btn_pe_read":       "Read & Preview",
		"btn_pe_convert":    "Convert to ZiiPOS",
		"col_itemcode":      "ItemCode",
		"col_name":          "Name",
		"col_price":         "Price",
		"col_category":      "Category",
		"col_menugroup":     "MenuGroup",
		"col_tax":           "Tax",
		"col_status":        "Status",
		"col_has_image":     "Image",
		"msg_done":          "Export completed!",
		"msg_err":           "Processing failed",
		"msg_no_file":       "Please select a file!",
		"msg_no_template":   "Template missing and download failed",
		"msg_reading":       "Reading file...",
		"msg_exporting":     "Exporting...",
		"msg_saving_images": "Saving images...",
		"msg_pe_done":       "PE Menu import completed!",
		"status_show":       "Show",
		"status_hide":       "Hide",
	},
	"cn": {
		"title":             "ZiiPOS 菜单转换器 V4.0 (Go)",
		"tab_excel":         "Excel 导入",
		"tab_pe":            "PE 菜单导入",
		"lbl_menu_file":     "菜单文件",
		"lbl_pe_file":       "PE 菜单文件",
		"lbl_output_folder": "输出文件夹",
		"btn_select":        "选择",
		"btn_browse":        "浏览",
		"btn_convert":       "转换",
		"btn_pe_read":       "读取预览",
		"btn_pe_convert":    "转换为 ZiiPOS",
		"col_itemcode":      "产品代码",
		"col_name":          "名称",
		"col_price":         "价格",
		"col_category":      "分类",
		"col_menugroup":     "菜单组",
		"col_tax":           "税率",
		"col_status":        "状态",
		"col_has_image":     "图片",
		"msg_done":          "导出完成！",
		"msg_err":           "处理失败",
		"msg_no_file":       "请选择文件！",
		"msg_no_template":   "模板缺失且下载失败",
		"msg_reading":       "正在读取文件...",
		"msg_exporting":     "正在导出...",
		"msg_saving_images": "正在保存图片...",
		"msg_pe_done":       "PE 菜单导入完成！",
		"status_show":       "展示",
		"status_hide":       "隐藏",
	},
	"jp": {
		"title":             "ZiiPOS メニュー変換 V4.0 (Go)",
		"tab_excel":         "Excel 取込",
		"tab_pe":            "PE メニュー取込",
		"lbl_menu_file":     "メニューファイル",
		"lbl_pe_file":       "PE メニューファイル",
		"lbl_output_folder": "出力フォルダ",
		"btn_select":        "選択",
		"btn_browse":        "参照",
		"btn_convert":       "変換",
		"btn_pe_read":       "読取・プレビュー",
		"btn_pe_convert":    "ZiiPOS に変換",
		"col_itemcode":      "商品コード",
		"col_name":          "名称",
		"col_price":         "価格",
		"col_category":      "カテゴリ",
		"col_menugroup":     "メニューグループ",
		"col_tax":           "税率",
		"col_status":        "ステータス",
		"col_has_image":     "画像",
		"msg_done":          "出力完了！",
		"msg_err":           "処理失敗",
		"msg_no_file":       "ファイルを選択してください！",
		"msg_no_template":   "テンプレート不足",
		"msg_reading":       "ファイル読込中...",
		"msg_exporting":     "出力中...",
		"msg_saving_images": "画像保存中...",
		"msg_pe_done":       "PE メニュー取込完了！",
		"status_show":       "表示",
		"status_hide":       "非表示",
	},
}

func T(key, lang string) string {
	if m, ok := strings[lang]; ok {
		if v, ok := m[key]; ok {
			return v
		}
	}
	if v, ok := strings["en"][key]; ok {
		return v
	}
	return key
}

func All(lang string) map[string]string {
	if m, ok := strings[lang]; ok {
		out := make(map[string]string, len(m))
		for k, v := range m {
			out[k] = v
		}
		return out
	}
	return strings["en"]
}
