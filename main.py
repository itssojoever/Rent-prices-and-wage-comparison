import xlwings as xl

filepath = "priceindexofprivaterentsukhistoricalseriesaccessible.xlsx"

wb1 = xl.Book(filepath)
sheet1 = wb1.sheets["Table 3"]

rentValues = sheet1.range("D5:D234").value
monthValues = sheet1.range("A5:A234").value

for value in rentValues:
    print (f"£{value}")


wb1.save(filepath)
wb1.close()

