import time
from requests_html import HTMLSession
import pandas as pd
import os
from openpyxl import load_workbook
from datetime import datetime

session = HTMLSession()

def get_profinance_rub_usd():
    r = session.get("https://www.profinance.ru/chart/usdrub/")
    rows = r.html.find('.stat.news > tr')
    for row in rows:
        selected_row = row.find('td')
        if (selected_row[0].text == "Курс доллара к рублю (USDRUB)"):
            ask = selected_row[2].text

    # print(float(ask))
    return float(ask)

def get_investing_brent_oil():
    r = session.get("https://ru.investing.com/commodities/brent-oil")

    brent_oil = r.html.find('[data-test="bid-value"] > span')[2].text

    # print(float(brent_oil.replace(',','.')))
    return float(brent_oil.replace(',','.'))

def write_to_excel(excel_name, _sheet_name, date_time):
    # print("date and time =", date_time)
    # print(os.path.exists(excel_name))

    df = pd.DataFrame({"USD": [get_profinance_rub_usd()],
                   "Brent-oil": [get_investing_brent_oil()],
                   "DateTime": [date_time]})

    if os.path.exists(excel_name):
        with pd.ExcelWriter(excel_name,mode="a",engine="openpyxl",if_sheet_exists="overlay") as writer:
            df.to_excel(writer, sheet_name=_sheet_name,header=None, startrow=writer.sheets[_sheet_name].max_row,index=False)
    else:
        writer = pd.ExcelWriter(excel_name, engine="openpyxl")
        df.to_excel(writer, index=False)
        writer.save()

def main(n, filename):
    if (filename == '' or filename.isspace()):
        filename = "result.xlsx"
    else:
        if not filename.endswith(".xlsx"):
            filename += ".xlsx"

    print("Results file: ", filename)
    print("Time between requests: ", n, " minutes")
    

    while (True):
        now = datetime.now()
        date_time = now.strftime("%d.%m.%Y %H:%M")

        print("Saving data...")
        write_to_excel(filename, "Sheet1", date_time)

        print("DateTime:", date_time)
        print("USD:", get_profinance_rub_usd())
        print("Brent:", get_investing_brent_oil())
        print("Save completed")
        time.sleep(n * 60)

print("Enter time(in minutes) between requests: ")
n = float(input())
while (n < 1):
    print("Time must be more than 1")
    n = float(input())

print("Enter a file name to save the results or press enter to apply the default value - 'results.xlsx'")
filename = input()

os.system("CLS")

main(n, filename)