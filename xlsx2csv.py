import openpyxl
import csv

if __name__ == "__main__":
    input_file = input("Enter input Excel file path (.xlsx): ").strip()
    output_file = input("Enter output CSV file path (.csv): ").strip()

    # Load the workbook and select the active sheet
    wb = openpyxl.load_workbook(input_file)
    sh = wb.active

    # Open a new CSV file in write mode
    with open(output_file, 'w', newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        # Iterate through rows and write to CSV
        for row in sh.iter_rows(values_only=True):
            writer.writerow(row)

    print("Conversion complete.")