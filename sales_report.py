import csv
import os

INPUT_FILE = "sales.csv"
OUTPUT_FILE = "sales_report.csv"

if not os.path.exists(INPUT_FILE):
    print("Error: sales.csv file not found!")
    exit()

total_sales = 0
total_quantity = 0
highest_sale = 0
best_product = ""

product_sales = {}

with open(INPUT_FILE, "r") as file:
    data = csv.DictReader(file)

    for item in data:
        product = item["product"]
        quantity = int(item["quantity"])
        price = float(item["price"])

        sales = quantity * price

        total_sales += sales
        total_quantity += quantity

        product_sales[product] = product_sales.get(product, 0) + sales

        if sales > highest_sale:
            highest_sale = sales
            best_product = product

if total_quantity > 0:
    average_sale = total_sales / total_quantity
else:
    average_sale = 0

with open(OUTPUT_FILE, "w", newline="") as file:
    writer = csv.writer(file)

    writer.writerow(["SALES ANALYSIS REPORT"])
    writer.writerow([])

    writer.writerow(["SUMMARY"])
    writer.writerow(["Total Quantity", total_quantity])
    writer.writerow(["Total Sales", round(total_sales, 2)])
    writer.writerow(["Best Product", best_product])
    writer.writerow(["Highest Sale", round(highest_sale, 2)])
    writer.writerow(["Average Sale", round(average_sale, 2)])

    writer.writerow([])
    writer.writerow(["PRODUCT-WISE SALES"])
    writer.writerow(["Product", "Sales"])

    for product, sales in product_sales.items():
        writer.writerow([product, round(sales, 2)])

print("================================")
print("     SALES ANALYSIS COMPLETE")
print("================================")
print("Total Quantity :", total_quantity)
print("Total Sales    :", round(total_sales, 2))
print("Best Product   :", best_product)
print("Highest Sale   :", round(highest_sale, 2))
print("Average Sale   :", round(average_sale, 2))
print("--------------------------------")
print("Report created:", OUTPUT_FILE)
print("================================")