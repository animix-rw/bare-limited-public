import _source

#Regular RS & Sales Daily file
daily_job = ["會員資料列表","項目業績報表","服務使用報表"]
for j in daily_job:
    _source.report_download(j, google_sheet_name="bare_maintenance")

_source.sales_rs_update()