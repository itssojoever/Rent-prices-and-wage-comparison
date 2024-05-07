import xlwings as xl
import matplotlib.pyplot as plt

filepath = "priceindexofprivaterentsukhistoricalseriesaccessible.xlsx"

wb1 = xl.Book(filepath)
sheet1 = wb1.sheets["Table 3"]

rentValues = sheet1.range("D5:D234").value
monthValues = sheet1.range("A5:A234").value

for value in rentValues:
    print (f"£{value}")


wb1.save(filepath)
wb1.close()

def generatePlot():

    plt.style.use("fivethirtyeight")
    fig, ax1 = plt.subplots()
    ax1.plot(monthValues, rentValues, label="Average rent price in England (£)", color="red")
    ax1.set_xlabel("Date")
    ax1.set_ylabel("Average rent in England (£)")

    plt.tight_layout()

    plt.show()

generatePlot()



