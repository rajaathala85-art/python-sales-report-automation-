import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def generate_report():

    file_path = filedialog.askopenfilename(
        title="Select Sales CSV File",
        filetypes=[("CSV Files", "*.csv")]
    )

    if not file_path:
        return

    try:
        total_sales = 0
        total_quantity = 0
        highest_sale = 0
        best_product = ""
        product_sales = {}
        product_quantity = {}

        with open(
            file_path,
            "r",
            newline="",
            encoding="utf-8-sig"
        ) as file:

            data = csv.DictReader(file)

            required_columns = {
                "product",
                "quantity",
                "price"
            }

            if not data.fieldnames:
                raise ValueError("CSV file is empty.")

            if not required_columns.issubset(set(data.fieldnames)):
                missing = required_columns - set(data.fieldnames)

                raise ValueError(
                    "Missing columns: "
                    + ", ".join(missing)
                )

            for item in data:

                product = item["product"].strip()
                quantity = int(item["quantity"])
                price = float(item["price"])

                if not product:
                    continue

                if quantity < 0 or price < 0:
                    raise ValueError(
                        f"Invalid value for {product}"
                    )

                sales = quantity * price

                total_sales += sales
                total_quantity += quantity

                product_sales[product] = (
                    product_sales.get(product, 0) + sales
                )

                product_quantity[product] = (
                    product_quantity.get(product, 0) + quantity
                )

                if sales > highest_sale:
                    highest_sale = sales
                    best_product = product

        if total_quantity > 0:
            average_sale = total_sales / total_quantity
        else:
            average_sale = 0

        # Ask where to save Excel report
        output_file = filedialog.asksaveasfilename(
            title="Save Excel Report",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Files", "*.xlsx")
            ],
            initialfile="sales_report.xlsx"
        )

        if not output_file:
            return

        # Create Excel workbook
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Sales Report"

        # Title
        sheet["A1"] = "SALES ANALYSIS REPORT"
        sheet["A1"].font = Font(
            bold=True,
            size=20
        )

        sheet.merge_cells("A1:D1")

        # Summary
        sheet["A3"] = "SUMMARY"
        sheet["A3"].font = Font(
            bold=True,
            size=14
        )

        summary = [
            ["Source File", os.path.basename(file_path)],
            ["Total Quantity", total_quantity],
            ["Total Sales", total_sales],
            ["Best Product", best_product],
            ["Highest Sale", highest_sale],
            ["Average Revenue Per Item", average_sale],
        ]

        row = 4

        for item in summary:
            sheet.cell(row=row, column=1).value = item[0]
            sheet.cell(row=row, column=2).value = item[1]
            row += 1

        # Product section
        start_row = 12

        sheet.cell(
            row=start_row,
            column=1
        ).value = "PRODUCT-WISE SALES"

        sheet.cell(
            row=start_row,
            column=1
        ).font = Font(
            bold=True,
            size=14
        )

        header_row = start_row + 1

        headers = [
            "Product",
            "Quantity",
            "Sales"
        ]

        for column, header in enumerate(headers, start=1):

            cell = sheet.cell(
                row=header_row,
                column=column
            )

            cell.value = header
            cell.font = Font(bold=True)
            cell.alignment = Alignment(
                horizontal="center"
            )

        current_row = header_row + 1

        for product in product_sales:

            sheet.cell(
                row=current_row,
                column=1
            ).value = product

            sheet.cell(
                row=current_row,
                column=2
            ).value = product_quantity[product]

            sheet.cell(
                row=current_row,
                column=3
            ).value = product_sales[product]

            current_row += 1

        # Number formatting
        sheet["B6"].number_format = '#,##0.00'
        sheet["B8"].number_format = '#,##0.00'
        sheet["B9"].number_format = '#,##0.00'

        for row_number in range(
            header_row + 1,
            current_row
        ):
            sheet.cell(
                row=row_number,
                column=3
            ).number_format = '#,##0.00'

        # Column widths
        widths = {
            "A": 28,
            "B": 20,
            "C": 20,
            "D": 20
        }

        for column, width in widths.items():
            sheet.column_dimensions[column].width = width

        # Save workbook
        workbook.save(output_file)

        result_text.set(
            f"Excel Report Generated!\n\n"
            f"Total Sales: ₹{total_sales:,.2f}\n"
            f"Best Product: {best_product}\n"
            f"Saved as: {os.path.basename(output_file)}"
        )

        messagebox.showinfo(
            "Success",
            "Professional Excel report created!"
        )

    except ValueError as error:

        messagebox.showerror(
            "Invalid CSV",
            str(error)
        )

    except Exception as error:

        messagebox.showerror(
            "Error",
            f"Something went wrong:\n\n{error}"
        )


# -----------------------------
# MAIN WINDOW
# -----------------------------

root = tk.Tk()

root.title("Sales Report Generator")
root.geometry("650x500")
root.resizable(False, False)

root.configure(bg="#f4f6f8")


header = tk.Frame(
    root,
    bg="#17202a",
    height=120
)

header.pack(fill="x")


title = tk.Label(
    header,
    text="Sales Report Generator",
    font=("Arial", 26, "bold"),
    fg="white",
    bg="#17202a"
)

title.pack(pady=(25, 5))


subtitle = tk.Label(
    header,
    text="CSV → Excel Business Automation",
    font=("Arial", 11),
    fg="#d5d8dc",
    bg="#17202a"
)

subtitle.pack()


content = tk.Frame(
    root,
    bg="#f4f6f8"
)

content.pack(
    fill="both",
    expand=True
)


instruction = tk.Label(
    content,
    text="Generate Professional Excel Report",
    font=("Arial", 18, "bold"),
    bg="#f4f6f8",
    fg="#17202a"
)

instruction.pack(pady=(45, 10))


info = tk.Label(
    content,
    text="Required: product • quantity • price",
    font=("Arial", 10),
    bg="#f4f6f8",
    fg="#566573"
)

info.pack(pady=5)


generate_button = tk.Button(
    content,
    text="SELECT CSV & CREATE EXCEL REPORT",
    font=("Arial", 12, "bold"),
    bg="#2874a6",
    fg="white",
    activebackground="#1f618d",
    activeforeground="white",
    padx=25,
    pady=14,
    relief="flat",
    cursor="hand2",
    command=generate_report
)

generate_button.pack(pady=30)


result_text = tk.StringVar()

result_label = tk.Label(
    content,
    textvariable=result_text,
    font=("Arial", 11),
    bg="white",
    fg="#17202a",
    width=65,
    height=5,
    relief="solid",
    bd=1
)

result_label.pack(pady=10)


footer = tk.Label(
    root,
    text="Python • CSV Automation • Excel Reporting",
    font=("Arial", 9),
    bg="#f4f6f8",
    fg="#7b7d7d"
)

footer.pack(
    side="bottom",
    pady=15
)


root.mainloop()