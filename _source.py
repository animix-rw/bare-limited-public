from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import json    
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
#import datetime
from datetime import datetime as dt
import shutil
import os
import os.path
import pandas as pd

# 1. Add an 'r' before the triple quotes to make it a Raw String


# Chrome Driver Options
options = Options()
options.add_argument("--disable-notifications")
options.add_argument('--no-sandbox')
options.add_argument('--verbose')
options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\tmp",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing_for_trusted_sources_enabled": False,
    "safebrowsing.enabled": False
})
options.add_argument('--disable-gpu')
options.add_argument('--disable-software-rasterizer')
#options.add_argument("blink-settings=imagesEnabled=false")
#options.add_argument("window-size=1920x1080")
options.add_argument("--start-maximized")
options.add_argument("--incognito")
options.add_argument('--headless')
options.add_argument("--window-size=2560,1440")


options.add_argument("--disable-dev-shm-usage")

chrome_path = os.getenv('CHROME_PATH')
chromedriver_path = os.getenv('CHROMEDRIVER_PATH')
options.binary_location = chrome_path
service = Service(executable_path=chromedriver_path)

#options.binary_location ='.//chrome-win64/chrome.exe'
#service = Service(executable_path=r'chromedriver-win64.exe')

absolute_path = os.path.dirname(os.path.abspath("__file__"))



# Google Sheets Creds
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
        "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds_dict = json.loads(os.environ["_GOOGLE_CRED"])
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)


__URL = os.environ['_URL']
__USERNAME = os.environ['_USERNAME']
__PASSWORD = os.environ['_PASSWORD']




def enable_download_headless(browser,relative_path):

    full_path = os.path.join(absolute_path, relative_path)

    browser.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd':'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': full_path}}
    browser.execute("send_command", params)


def report_download(report_name, google_sheet_name, date_from=None):

    ##driver = webdriver.Chrome(PATH, options=options)
    driver = webdriver.Chrome(service=service, options=options)
    enable_download_headless(driver, "excel")

    driver.get(__URL)
    driver.switch_to.frame("mainFrame");
    driver.switch_to.frame("SYSTEM");

    driver.set_window_size(2560, 1440)
    driver.execute_script("document.body.style.zoom='80%'")

    log_name = driver.find_element("id","USERNAME")
    log_pw = driver.find_element("id","PASSWORD")
    log_name.send_keys(__USERNAME)
    log_pw.send_keys(__PASSWORD, Keys.RETURN)

    driver.switch_to.default_content();
    driver.switch_to.frame("mainFrame");
    driver.switch_to.frame("SYSTEM");

    wait = WebDriverWait(driver, 3)

    element = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "報表")))
    element.click()
    #driver.find_element(By.LINK_TEXT,"報表").click()
 
    element2 = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, report_name)))
    element2.click()
    #driver.find_element(By.LINK_TEXT, report_name).click()
    
    time.sleep(10)
    try:
        shop = Select(driver.find_element("id","BRANCH"))
        shop.select_by_visible_text("(以下所有)")
    except:
        pass

    #time.sleep(10)
    #Only apply to full year extraction in bare_sales
    #if google_sheet_name=="bare_sales" and (report_name=="項目業績報表" or report_name=="新舊會員營業額報表"):
    #    driver.find_element("id","DATE_FROM").send_keys(Keys.CONTROL + 'A', date_from)
    #    driver.find_element("id","DATE_FROM").send_keys(date_from)

    time.sleep(10)
    if report_name=="套票結餘報表":
        Select(driver.find_element("id","EXPIRE")).select_by_visible_text("未過期")

    ActionChains(driver).key_down(Keys.CONTROL).send_keys('+').key_up(Keys.CONTROL).perform()

    wait.until(EC.element_to_be_clickable((By.ID, "BUTTON_EXPORT")))
    driver.find_element(By.ID,"BUTTON_EXPORT").click()

    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "是")))
    driver.find_element(By.LINK_TEXT,"是").click()

    driver.implicitly_wait(20)
    
    # Wait at most 20 second until file available
    start_time = time.time()
    path = os.path.join(absolute_path, 'excel', f"{report_name}.csv")
    [time.sleep(1) for _ in iter(lambda: not os.path.exists(path) and time.time() - start_time < 10, False)]
    
    driver.close()
    driver.quit()

        
    _df = pd.read_csv(path)
    
    if report_name=="項目業績報表":
        _df['內部備註'] = _df['內部備註'].replace('', float('nan'))
        _df['內部備註'] = _df['內部備註'].fillna('NUL')
        
    #_df.fillna('', inplace=True)
    _df.fillna(0, inplace=True)

    sheet_ = client.open(google_sheet_name).worksheet(report_name)
    sheet_.clear()
    sheet_.update([_df.columns.values.tolist()] + _df.values.tolist())

    os.remove(path)
    print('Finished',report_name)







def sales_analysis_data():
    '''
    Input data
    Output two dataframe for Looker
    Working for sales analysis
    '''

        # Convert to Dataframe
    def convert_df(input_list):
        df = pd.DataFrame(input_list[1:],columns=input_list[0])
        return df


    # Retrieve related document
    sales_previous = client.open("bare_sales").worksheet("項目業績報表_previous").get_all_values()
    sales_current = client.open("bare_sales").worksheet("項目業績報表").get_all_values()
    oldnew_previous = client.open("bare_sales").worksheet("新舊會員營業額報表_previous").get_all_values()
    oldnew_current = client.open("bare_sales").worksheet("新舊會員營業額報表").get_all_values()
    txcate = client.open("bare_sales").worksheet("套票資料列表").get_all_values()
    mbrlist = client.open("bare_maintenance").worksheet("會員資料列表").get_all_values()
    onsales = [row[:4] for row in client.open("bare_maintenance").worksheet("name").get_all_values()]

    #Append both previous and current sales
    #sales_total = sales_previous + sales_current[1:]
    #oldnew_total = oldnew_previous + oldnew_current[1:]

    #sales_total = convert_df(sales_total)
    #oldnew_total = convert_df(oldnew_total)

    sales_previous = convert_df(sales_previous)
    sales_current = convert_df(sales_current)
    oldnew_previous = convert_df(oldnew_previous)
    oldnew_current = convert_df(oldnew_current)

    def update_latest_record(previous_record, latest_record, date):
        """Input two dataframes, check if previous record
        contain same date as latest record. If so, replace it with latest record
        and merge them"""
        # Convert the 'date' columns to datetime 
        previous_record[date] = pd.to_datetime(previous_record[date]) 
        latest_record[date] = pd.to_datetime(latest_record[date])

        # Find the dates in latest_record 
        latest_dates = latest_record[date].dt.date.unique()

        # Remove records in previous_record with dates in latest_record 
        previous_record = previous_record[~previous_record[date].dt.date.isin(latest_dates)]

        # Append latest_record to previous_record 
        updated_record = pd.concat([previous_record, latest_record])
        updated_record[date] = updated_record[date].astype(str)

        return updated_record

    sales_total = update_latest_record(sales_previous, sales_current, '單據日期')
    oldnew_total = update_latest_record(oldnew_previous, oldnew_current, '發票日期')

    txcate = convert_df(txcate)
    mbrlist = convert_df(mbrlist)
    onsales = convert_df(onsales)
    onsales.rename(columns={'銷售員名稱':'銷售員名稱(On Sales)', '分店':'分店(On Sales)'}, inplace=True)


    # Left join tx cate to gain main tx category
    sales_output = sales_total.merge(txcate[['套票編號','主類別','次類別']], left_on='項目編號', right_on='套票編號', how='left')#.drop(columns=['套票編號'])
    sales_output['主類別'] = sales_output['主類別'].fillna('其他')
    sales_output['次類別'] = sales_output['次類別'].fillna('其他')
    # Left join old new total to gain mbr code, cate
    sales_output = sales_output.merge(oldnew_total[['發票編號','會員編號','營業額性質']], left_on='單據編號', right_on='發票編號', how='left')
    # Left join mbr list to gain follow up therapist
    sales_output = sales_output.merge(mbrlist[['會員編號','跟進員工']], left_on='會員編號', right_on='會員編號', how='left')
    # Left join on sales to find on sales therapist
    sales_output = sales_output.merge(onsales[["銷售員名稱(On Sales)","分店(On Sales)"]],left_on="銷售員名稱", right_on="銷售員名稱(On Sales)", how="left")


    # Drop the extra columns
    sales_output.drop(columns=['套票編號','發票編號'], axis=1)
    # If not on sales, the public
    sales_output[['銷售員名稱(On Sales)','分店(On Sales)']] = sales_output[['銷售員名稱(On Sales)','分店(On Sales)']].fillna('Public')
    #Data Format
    sales_output['單據日期'] = pd.to_datetime(sales_output['單據日期'])
    sales_output['業績'] = sales_output['業績'].astype('float')


    # Group the total sales to be joined by mbr
    group_sales = sales_output.groupby("會員編號").agg(
        First_Purchase = pd.NamedAgg(column="單據日期", aggfunc="min"),
        Last_Purchase = pd.NamedAgg(column="單據日期", aggfunc="max"),
        Total_Spending = pd.NamedAgg(column="業績", aggfunc='sum'))

    # Pivot for the main tx category
    group_txpen = sales_output.groupby(['會員編號','主類別']).agg(
        Treatment_Purchase = pd.NamedAgg(column="業績", aggfunc="sum"))

    group_txpen_pivot = group_txpen.pivot_table(index='會員編號', columns='主類別', values='Treatment_Purchase',fill_value=0)



    # Prepare total member base list
    member_output = mbrlist.drop(columns=['加入分店','證件號碼','會員類別','性別','手提','電郵','地址','職業','居住地區','來源','介紹會員編號','介紹會員名稱','學歷','收入'])

    # Left Join on sales
    member_output = member_output.merge(onsales[["銷售員名稱(On Sales)","分店(On Sales)"]],left_on="跟進員工", right_on="銷售員名稱(On Sales)", how="left")
    member_output[['銷售員名稱(On Sales)','分店(On Sales)']] = member_output[['銷售員名稱(On Sales)','分店(On Sales)']].fillna('Public')
    # Left Join sales fig
    member_output = member_output.merge(group_sales, left_on="會員編號", right_on="會員編號", how="left")
    # Left join main tx category
    member_output = pd.merge(member_output, group_txpen_pivot, left_on='會員編號', right_on='會員編號', how='left')


    # Upload to google sheet
    sales_output[['單據日期','業績','內部備註']] = sales_output[['單據日期','業績','內部備註']].astype('str')
    sales_output.fillna(0, inplace=True)
    sales_output['內部備註'] = sales_output['內部備註'].replace(0,'NUL')
    _ = client.open('bare_sales').worksheet('sales_output')
    _.clear()
    _.update([sales_output.columns.values.tolist()] + sales_output.values.tolist())

    member_output = member_output.astype('str')
    member_output.fillna(0, inplace=True)
    _ = client.open('bare_sales').worksheet('member_output')
    _.clear()
    _.update([member_output.columns.values.tolist()] + member_output.values.tolist())

    print("Uploaded Sales Analysis Data Success")


def sales_rs_update():
    '''
    Input data
    Output two dataframe for Looker
    Working for sales analysis
    '''

    import gspread
    from oauth2client.service_account import ServiceAccountCredentials



    # retrieve raw relevant document
    sales_hourly = client.open("bare_maintenance").worksheet("項目業績報表").get_all_values()
    # split invoice calculation in python, not google
    #sales_adjust = client.open("bare_maintenance").worksheet("adjust_sales").get_all_values()

    name_ = client.open("bare_maintenance").worksheet("name").get_all_values()
    outstanding_course = client.open("bare_maintenance").worksheet("套票結餘報表").get_all_values()
    member_detail = client.open("bare_maintenance").worksheet("會員資料列表").get_all_values()
    course_useage = client.open("bare_maintenance").worksheet("服務使用報表").get_all_values()

    # Convert to Dataframe
    def convert_df(input_list):
        df = pd.DataFrame(input_list[1:],columns=input_list[0])
        return df

    # Output raw dataframe for further processing
    sales_hourly_adjust = sales_hourly #+ sales_adjust[1:]
    sales_hourly_adjust = convert_df(sales_hourly_adjust)

    # split invoice calculation in phython, not google
    def adjust_sales_df(sales_hourly_list):
        df_sales_hourly = convert_df(sales_hourly_list)
        
        # Filter split invoice rows
        filtered_df = df_sales_hourly[
            (df_sales_hourly['內部備註'].str.contains('&', na=False)) &
            (~df_sales_hourly['內部備註'].str.startswith('&', na=False)) &
            (df_sales_hourly['業績'].astype(str) != '0')
        ].copy()
    
        # Apply transformations for each filtered row
        new_rows = []
        for index, row in filtered_df.iterrows():
          # Step 1: Divide "業績" by -2 for the original row
          original_row = row.copy()
          original_row['業績'] = float(original_row['業績']) / -2
          new_rows.append(original_row)
    
          # Step 2: Replace "銷售員名稱" and divide "業績" by 2
          new_row = row.copy()
          internal_remark = new_row['內部備註']
          if '&' in internal_remark:
            salesperson_name = internal_remark.split('&')[0]#.strip()
            new_row['銷售員名稱'] = salesperson_name
            new_row['業績'] = float(new_row['業績']) / 2
            new_rows.append(new_row)
    
        # Create a new DataFrame from the processed rows
        processed_df = pd.DataFrame(new_rows)
    
        # upload to Google sheet
        sheet = client.open('bare_maintenance').worksheet('adjust_sales_')
        sheet.clear()
        sheet.update([processed_df.columns.values.tolist()] + processed_df.values.tolist())
    
        return processed_df
    split_sales_data = adjust_sales_df(sales_hourly)
    
    # append split invoice data to original data
    sales_hourly_adjust = pd.concat([sales_hourly_adjust, split_sales_data], ignore_index=True)

    
    # remove useless header and convert to values
    sales_hourly_adjust = sales_hourly_adjust[sales_hourly_adjust['單據日期'] != '單據日期']
    #sales_hourly_adjust['業績'] = sales_hourly_adjust['業績'].str.replace(',', '')
    sales_hourly_adjust['業績'] = sales_hourly_adjust['業績'].astype(float)

    name_ = convert_df(name_)
    outstanding_course = convert_df(outstanding_course)
    member_detail = convert_df(member_detail)
    course_useage = convert_df(course_useage)


    # Create master sheet
    sole_name_ = name_["銷售員名稱"][name_["銷售員名稱"] != ""]
    sole_date_ = name_["日期"][name_["日期"] != ""]

    master_ = pd.merge(sole_name_, sole_date_, how="cross")
    master_ = master_.merge(name_[["銷售員名稱","銷售員編號","分店","銷售員當月銷售目標"]], left_on="銷售員名稱", right_on="銷售員名稱", how="left")
    master_ = master_.merge(name_[["日期","工作日","全月工作日","Is_WD"]], left_on="日期", right_on="日期", how="left")
    # Convert to value
    master_[['工作日', '全月工作日', '銷售員當月銷售目標','Is_WD']] = master_[['工作日', '全月工作日', '銷售員當月銷售目標','Is_WD']].replace('', float('nan'))
    master_[['工作日', '全月工作日', '銷售員當月銷售目標','Is_WD']] = master_[['工作日', '全月工作日', '銷售員當月銷售目標','Is_WD']].fillna(0)
    master_[['工作日', '全月工作日', '銷售員當月銷售目標','Is_WD']] = master_[['工作日', '全月工作日', '銷售員當月銷售目標','Is_WD']].astype(float)


    # Each therapist daily sales
    # Join on sales est
    sole_name_ = sole_name_.to_frame(name='銷售員名稱')
    sales_hourly_adjust.rename(columns={'銷售員名稱':'銷售員名稱(POS)'}, inplace=True)
    sales_hourly_adjust = sales_hourly_adjust.merge(sole_name_[['銷售員名稱']], left_on='銷售員名稱(POS)', right_on='銷售員名稱', how='left')
    sales_hourly_adjust['銷售員名稱'] = sales_hourly_adjust['銷售員名稱'].fillna('Public')
    # Insert new code
    group_thp = sales_hourly_adjust.groupby(['銷售員名稱', '單據日期'])['業績'].sum().reset_index(name='DailySales')


    # Join therapist daily sales to master
    master_ = master_.merge(group_thp[["銷售員名稱","單據日期","DailySales"]], left_on=["銷售員名稱","日期"], right_on=["銷售員名稱","單據日期"], how="left")
    # Drop useless date filed after join
    master_.drop(columns=['單據日期'], inplace=True)
    master_['DailyTargetSales'] = master_['工作日'] / master_['全月工作日'] * master_['銷售員當月銷售目標']
    # Compute the daily target, not acc daily target
    master_['DailyTargetSales'] = master_.apply(
        lambda row: 1 / row['全月工作日'] * row['銷售員當月銷售目標'] if row['Is_WD'] == 1 else 0,
        axis=1)


    # Prepare RS target (update monthly at first day of month)
    outstanding_course[['套票值','已用套票值','剩餘套票值','金額','已用金額','剩餘金額']] = outstanding_course[['套票值','已用套票值','剩餘套票值','金額','已用金額','剩餘金額']].astype(float)
    outstanding_course = outstanding_course[outstanding_course['剩餘金額']>0]

    # Compute outsatnding for individual customer
    outstanding_course = outstanding_course.groupby(["會員編號"]).agg(
        套票值 = pd.NamedAgg(column="套票值", aggfunc="sum"),
        已用套票值 = pd.NamedAgg(column="已用套票值", aggfunc="sum"),
        剩餘套票值 = pd.NamedAgg(column="剩餘套票值", aggfunc='sum'),
        金額 = pd.NamedAgg(column="金額", aggfunc="sum"),
        已用金額 = pd.NamedAgg(column="已用金額", aggfunc="sum"),
        剩餘金額 = pd.NamedAgg(column="剩餘金額", aggfunc='sum'))

    # Find the following staff
    outstanding_course = outstanding_course.merge(member_detail[['會員編號','會員名稱','跟進員工']], left_on='會員編號', right_on='會員編號', how='left')

    # Joined original on sales therapist
    outstanding_course = outstanding_course.merge(sole_name_[['銷售員名稱']], left_on='跟進員工', right_on='銷售員名稱', how='left')
    # Define public therapist
    outstanding_course['銷售員名稱'] = outstanding_course['銷售員名稱'].fillna('Public')

    # Join last monthly realized date into the outstanding master
    # Compute Last Realized date
    course_useage_last = course_useage.groupby(['會員編號']).agg(
        LastRealizedDate = pd.NamedAgg(column="單據日期", aggfunc="max"),
        MonthlyRSAmt = pd.NamedAgg(column="使用金額", aggfunc="sum"),
        MonthlyRSQty = pd.NamedAgg(column="使用值", aggfunc="sum"))
    course_useage_last = course_useage_last.reset_index()
    # Join to RS master
    outstanding_course = outstanding_course.merge(course_useage_last[['會員編號','LastRealizedDate']], left_on='會員編號', right_on='會員編號', how='left')
    
    # Insert the current report month
    outstanding_course["Report_Month"] = dt.now().strftime("%Y-%m")

    # Full month RS Target
    #group_outstanding_course = outstanding_course.groupby(['銷售員名稱']).agg(
    #    Full_RS_Target = pd.NamedAgg(column="會員編號", aggfunc="count"))
    group_outstanding_course = outstanding_course.groupby(['銷售員名稱'])['會員編號'].count().reset_index(name='Full_RS_Target')



    ## Daily Actual RS

    # Compute from course usage, may end up in multiple count of one customer.
    # Join with corresponding therapist
    # course_useage = course_useage.merge(member_detail[['會員編號','跟進員工']], left_on='會員編號', right_on='會員編號', how='left')
    # Join on sales list
    # course_useage = course_useage.merge(sole_name_[['銷售員名稱']], left_on='跟進員工', right_on='銷售員名稱', how='left')
    # Define public therapist
    # course_useage['銷售員名稱'] = course_useage['銷售員名稱'].fillna('Public')

    #group_course_usage = course_useage.groupby(['單據日期','銷售員名稱']).agg(
    #    DailyRS = pd.NamedAgg(column="會員編號", aggfunc="count"))
    # group_course_usage = course_useage.groupby(['單據日期','銷售員名稱'])['會員編號'].count().reset_index(name='DailyRS')

    # Compute from RS target as base.
    group_course_usage = outstanding_course.groupby(['銷售員名稱','LastRealizedDate']).agg(
        DailyRS = pd.NamedAgg(column="會員編號", aggfunc="count")
    )
    group_course_usage = group_course_usage.reset_index()


    # Join RS related data into master_

    # Full month target
    master_ = master_.merge(group_outstanding_course[['銷售員名稱','Full_RS_Target']], left_on='銷售員名稱', right_on='銷售員名稱', how='left')
    # Compute the daily target, not acc daily target
    master_['DailyTargetRS'] = master_.apply(
        lambda row: 1 / row['全月工作日'] * row['Full_RS_Target'] if row['Is_WD'] == 1 else 0,
        axis=1)

    # Daily RS
    master_ = master_.merge(group_course_usage[['LastRealizedDate','銷售員名稱','DailyRS']], left_on=['日期','銷售員名稱'], right_on=['LastRealizedDate','銷售員名稱'], how='left')
    # Drop useless date filed after join
    master_.drop(columns=['LastRealizedDate'], inplace=True)

    # Replace NaN to 0
    master_['DailyRS'] = master_['DailyRS'].fillna(0)
    master_['DailySales'] = master_['DailySales'].fillna(0)


    #Output file to google sheets
    #master_ == daily_master
    #outstanding_course == rs_master
    #sales_hourly_adjust == sales_master


    #_ = client.open('bare_maintenance').worksheet('daily_master')
    #master_.fillna(0, inplace=True)
    #_.clear()
    #_.update([master_.columns.values.tolist()] + master_.values.tolist())

    #outstanding_course.fillna(0, inplace=True)
    #_ = client.open('bare_maintenance').worksheet('rs_master')
    #_.clear()
    #_.update([outstanding_course.columns.values.tolist()] + outstanding_course.values.tolist())

    #sales_hourly_adjust.fillna(0, inplace=True)
    #_ = client.open('bare_maintenance').worksheet('sales_master')
    #_.clear()
    #_.update([sales_hourly_adjust.columns.values.tolist()] + sales_hourly_adjust.values.tolist())

    #print("Finish updating sales and RS data") # master_.fillna(0, inplace=True)


    def append_google_record(google_wb, google_ws, df_new_data, data_Date_name):

        # Open the Google Sheet
        sheet = client.open(google_wb).worksheet(google_ws)
        records = sheet.get_all_records()
        df_google_sheet = pd.DataFrame(records)

        # New data replace NA
        df_new_data.fillna(0, inplace=True)
        df_new_data[data_Date_name] = pd.to_datetime(df_new_data[data_Date_name])

        df_google_sheet[data_Date_name] = pd.to_datetime(df_google_sheet[data_Date_name])

        # Filter the Google Sheet DataFrame to exclude records with dates in the new data
        new_dates = df_new_data[data_Date_name].dt.date.unique()
        df_filtered = df_google_sheet[~df_google_sheet[data_Date_name].dt.date.isin(new_dates)]

        # Append the new data to the filtered DataFrame
        df_combined = pd.concat([df_filtered, df_new_data])
        df_combined[data_Date_name] = df_combined[data_Date_name].astype(str)

        # Clear the Google Sheet
        sheet.clear()

        # Update the Google Sheet with the combined DataFrame
        sheet.update([df_combined.columns.values.tolist()] + df_combined.values.tolist())

        print(f"{google_ws} updated successfully!")

    append_google_record('bare_maintenance', 'daily_master', master_, '日期')
    append_google_record('bare_maintenance', 'rs_master', outstanding_course, 'Report_Month')
    append_google_record('bare_maintenance', 'sales_master', sales_hourly_adjust, '單據日期')