import _source

#Sales analysis
#_source.report_download("項目業績報表",google_sheet_name="bare_sales",date_from="2024-01-01")
_source.report_download("項目業績報表",google_sheet_name="bare_sales")
_source.report_download("新舊會員營業額報表",google_sheet_name="bare_sales")
_source.report_download("套票資料列表",google_sheet_name="bare_sales")
_source.sales_analysis_data()