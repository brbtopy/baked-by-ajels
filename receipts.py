from reportlab.pdfgen import canvas
from reportlab.lib import colors
import os
import requests
import json
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import date
from stat import S_IREAD, S_IRGRP, S_IROTH
from stat import S_IWUSR
from num2words import num2words
from PIL import Image, ImageEnhance
import fitz  # PyMuPDF
import io


def read_counter(directory):
    called = True
    if called:
        count_file = open(directory, "r")
        count = count_file.read()
        count_file.close()
    return count

def increase_counter(directory):
    called = True
    if called:
        count_file = open(directory, "r")
        count = count_file.read()
        count_file.close()

        count_file = open(directory, "w")
        count = int(count) + 1
        count_file.write(str(count))
        count_file.close()
    return count

def drawMyRuler(pdf):
    pdf.drawString(100,310, 'x100')
    pdf.drawString(200,310, 'x200')
    pdf.drawString(300,310, 'x300')
    pdf.drawString(400,310, 'x400')
    pdf.drawString(500,310, 'x500')
    pdf.drawString(600,310, 'x600')
    pdf.drawString(700,310, 'x700')

    pdf.drawString(10,100, 'y100')
    pdf.drawString(10,200, 'y200')
    pdf.drawString(10,300, 'y300')
    pdf.drawString(10,400, 'y400')
    pdf.drawString(10,500, 'y500')
    pdf.drawString(10,600, 'y600')
    pdf.drawString(10,700, 'y700')
    pdf.drawString(10,800, 'y800')


excel_file = "C:\\Users\\samas\\OneDrive\\Desktop\\proj\\ajels\\baked-by-ajels-receipt\\baked by ajels order workbook.xlsx"
os.chmod(excel_file, S_IWUSR|S_IREAD)

workbook = load_workbook(filename=excel_file)
main_sheet = workbook["main"]

main_counter_file = "C:\\Users\\samas\\OneDrive\\Desktop\\proj\\ajels\\baked-by-ajels-receipt\\counter\\counter.txt"
receipt_counter_file =  "C:\\Users\\samas\\OneDrive\\Desktop\\proj\\ajels\\baked-by-ajels-receipt\\receipt_docs\\counter\\count.txt"

os.chmod(receipt_counter_file, S_IWUSR|S_IREAD)
os.chmod(main_counter_file, S_IWUSR|S_IREAD)

main_counter = read_counter(main_counter_file)
receipt_no = read_counter(receipt_counter_file)

today = date.today()

print("[+] Welcome to Baked By Ajels Receipt Generator [+]")
print("[+] Please provide the following details to generate a receipt [+]")
print("[+] Please ensure that the amount is in Ghana Cedis [+]")

username = input("Who is placing the order? ")
username = username.title()

products = str(input(f"What product(s) did {username} purchase? "))
products = products.title()
new_prod = products.replace("Of", "of")
new_prod = new_prod.replace("And", "and")

amount = int(input(f"How much is {username} paying? "))
amount_dsp = ("%.2f" % amount)

print("[+] Data collected. Generating receipt and updating records [+]")

pdf = canvas.Canvas("C:\\Users\\samas\\OneDrive\\Desktop\\proj\\ajels\\baked-by-ajels-receipt\\receipt_docs\\example.pdf")

pdf.setTitle(username + " Baked By Ajels Receipt")
pdf.setPageSize((800, 300))
pdf.setLineWidth(0)

pdf.setFont("Courier-Bold", 30)
pdf.drawString(50, 250, "RECEIPT")

#set receipt display details 
pdf.setFont("Courier", 14)
pdf.drawString(580, 245, "Receipt Number:  " + str(receipt_no))
date_format = today.strftime("%d-%b-%Y")
pdf.drawString(580, 215, "Date: " + date_format)

#set user details
pdf.setFont("Helvetica", 12)
pdf.drawString(50, 170, "Received from")
pdf.setFont("Helvetica", 14)
pdf.drawString(186, 170, username)
pdf.line(140, 165, 370, 165)

pdf.setFont("Helvetica", 12)
pdf.drawString(390, 170, "the amount in GHC")
pdf.setFont("Helvetica", 14)
pdf.drawString(525, 170, str(amount_dsp))
pdf.line(510, 165, 590, 165)

pdf.setFont("Helvetica", 12)
pdf.drawString(50, 130, "Amount in words")
num = amount 
num_words = num2words(num)
num_words = num_words.title()
if amount == 1:
  num_words = num_words + " Ghana Cedi only"
if amount > 1:
  num_words = num_words + " Ghana Cedis only"
pdf.setFont("Helvetica", 14)
pdf.drawString(170, 130, num_words)
pdf.line(150, 125, 590, 125)

pdf.setFont("Helvetica", 12)
pdf.drawString(50, 90, "For")
pdf.line(80, 85, 590, 85)
pdf.setFont("Helvetica", 14)
pdf.drawString(95, 90, new_prod)

#set up recipient tagging
pdf.setFont("Helvetica", 11)
pdf.drawString(50, 40, "Recipient :")
pdf.setFont("Helvetica", 13)
pdf.drawString(120, 40, "Baked by Ajels")
pdf.line(110, 35, 220, 35)

#set up thank you tagging
thanks = os.path.join(os.getcwd(), "C:\\Users\\samas\\OneDrive\\Desktop\\proj\\ajels\\baked-by-ajels-receipt\\pictures\\thanks.png")
pdf.drawImage(thanks, 633, 30, width=98, height=90)

#set footer
pdf.setStrokeColor(colors.gold)
pdf.setFillColor(colors.gold)
pdf.rect(0, 0, width=800, height=22, stroke=1, fill=1)
pdf.setFillColor(colors.black)
pdf.setFont("Helvetica", 12)
pdf.drawString(150, 7, "Call: +233 54 056 3300  |  WhatsApp: +233 20 952 7289  |  Instagram: @Baked_by_Ajels")
pdf.save()

#################################################################################
#appending values
cell_number = read_counter(main_counter_file)

main_sheet["A" + str(cell_number)] = username.title()

main_sheet["B" + str(cell_number)] = int(amount)
main_sheet["B" + str(cell_number)].number_format = '0.00'

date_format = today.strftime("%B %d, %Y")
main_sheet["C" + str(cell_number)] = date_format
main_sheet["C" + str(cell_number)].alignment = Alignment(horizontal='left', vertical='bottom')

main_sheet["D" + str(cell_number)] = new_prod


workbook.save(filename=excel_file)
workbook.close()

#################################################################################
# Adding watermark to receipt
document_path = 'C:\\Users\\samas\\OneDrive\\Desktop\\proj\\ajels\\baked-by-ajels-receipt\\receipt_docs\\example.pdf'
watermark_path = 'C:\\Users\\samas\\OneDrive\\Desktop\\proj\\ajels\\baked-by-ajels-receipt\\pictures\\logo.png'
opacity = 0.07

# Open the original document
pdf_document = fitz.open(document_path)
page = pdf_document.load_page(0)  # Load the first page

# Increase the DPI for better clarity
zoom = 2  # Increase the zoom factor to improve resolution
mat = fitz.Matrix(zoom, zoom)
pix = page.get_pixmap(matrix=mat)

document = Image.open(io.BytesIO(pix.tobytes()))

# Resize the watermark to a larger size
document_width, document_height = document.size
watermark = Image.open(watermark_path).convert("RGBA")
watermark_width = document_width // 3
watermark_height = int(watermark.size[1] * (watermark_width / watermark.size[0]))
watermark = watermark.resize((watermark_width, watermark_height), Image.LANCZOS)

# Adjust the opacity of the watermark
alpha = watermark.split()[3]
alpha = ImageEnhance.Brightness(alpha).enhance(opacity)
watermark.putalpha(alpha)

# Calculate the position to center the watermark
x = (document_width - watermark_width) // 2
y = (document_height - watermark_height) // 2

# Paste the watermark onto the document
document.paste(watermark, (x, y), watermark)

# Save the result
output_file = f"C:\\Users\\samas\\OneDrive\\Desktop\\proj\\ajels\\baked-by-ajels-receipt\\receipt_docs\\{username}'s Receipt.pdf"
document.save(output_file, format="PDF")
pdf_document.close()

os.remove("C:\\Users\\samas\\OneDrive\\Desktop\\proj\\ajels\\baked-by-ajels-receipt\\receipt_docs\\example.pdf")

#################################################################################
#increase counters and set read-only on files
increase_counter(receipt_counter_file)
increase_counter(main_counter_file)
os.chmod(receipt_counter_file, S_IREAD|S_IRGRP|S_IROTH)
os.chmod(main_counter_file, S_IREAD|S_IRGRP|S_IROTH)
os.chmod(excel_file, S_IREAD|S_IRGRP|S_IROTH)

print("[+] Receipt generated successfully [+]")
print("[+] Receipt saved in the receipt_docs folder [+]")