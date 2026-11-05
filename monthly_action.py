import _source

# Prepare RS target at first day of month
daily_job = ["套票結餘報表"]
for j in daily_job:
    _source.report_download(j, google_sheet_name="bare_maintenance")