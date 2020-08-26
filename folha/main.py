import subprocess
from scripts.generation import generate_csv
from scripts.conversion import csv_to_xlsx, xlsx_to_pdf

if __name__ == "__main__":
    csv = generate_csv()
    xlsx = csv_to_xlsx(csv)
