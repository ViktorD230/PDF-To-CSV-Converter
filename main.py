import os
import tkinter as tk
from tkinter import filedialog
import csv
import PyPDF2
import pytesseract
from PIL import Image

print("[-+-] starting pdf_csv.py...")
print("[-+-] import a pdf and convert it to a csv")

# -----------------------------------------------------------------------------

def pdf_csv():
    print("[-+-] default filenames:")
    filename = "sample1"
    csv = filename + ".csv"
    print(csv + "\n")

    print("[-+-] default directory:")
    print("[-+-] (based on the current working directory of the python file)")

    defaultdir = os.getcwd()
    print(defaultdir + "\n")

    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Open the file dialog to choose a PDF file
    pdf_path = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if pdf_path:
        print("[-+-] Selected PDF file:", pdf_path)
        pdf_flag = True
    else:
        print("[-+-] No PDF file selected.")
        pdf_flag = False

    if pdf_flag:
        pdf_filename = os.path.basename(pdf_path)
        csv = os.path.splitext(pdf_filename)[0] + ".csv"
        csv_path = os.path.join(defaultdir, csv)

        try:
            print("[-+-] looking for default csv...")
            open(csv_path, "r")
            print("[-+-] csv found: " + csv + "\n")
        except IOError:
            print("[-+-] did not find csv at default file path!")
            print("[-+-] creating a blank csv file: " + csv + "... \n")
            open(csv_path, "w")

        print("[-+-] converting pdf to csv...")
        try:
            pdf_data = extract_text_from_pdf(pdf_path)
            save_text_as_csv(pdf_data, csv_path)
            print("[-+-] pdf to csv conversion complete!\n")
        except Exception as e:
            print("[-+-] pdf to csv conversion failed!")
            print("[-+-] Error:", e)
            print("[-+-] converted csv file can be found here: " + csv_path + "\n")
    print("[-+-] finished pdf_csv.py successfully!")

def extract_text_from_pdf(pdf_path):
    pdf_data = ""
    with open(pdf_path, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfFileReader(pdf_file)
        for page_num in range(pdf_reader.numPages):
            page = pdf_reader.getPage(page_num)
            pdf_data += extract_text_from_image(page)
    return pdf_data

def extract_text_from_image(page):
    image = page.toImage()
    image_path = "temp_image.png"
    image.save(image_path)
    text = pytesseract.image_to_string(Image.open(image_path), lang="eng")
    os.remove(image_path)
    return text

def save_text_as_csv(pdf_data, csv_path):
    with open(csv_path, "w", newline='', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)
        lines = pdf_data.split('\n')
        for line in lines:
            csv_writer.writerow([line])

# Call the function to execute the script
pdf_csv()
