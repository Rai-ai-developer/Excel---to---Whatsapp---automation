import openpyxl
import pywhatkit
import time

# Excel file load
wb = openpyxl.load_workbook('contacts.xlsx')
sheet = wb.active

print("Starting Automation...")

for row in sheet.iter_rows(min_row=2, values_only=True):
    phone = str(row[0])
    msg = str(row[1])
    try:
        pywhatkit.sendwhatmsg_instantly(phone, msg)
        print(f"Sent to {phone}")
        time.sleep(10)
    except Exception as e:
        print(f"Error {phone}: {e}")

print("Done!")
