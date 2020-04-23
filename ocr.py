from PIL import Image
import pytesseract
from pdf2image import convert_from_path
import os
from openpyxl import load_workbook

wb = load_workbook('/Users/jeffmajnaric/Google Drive/Personal/Finance/RSE_Paystubs.xlsx')
sheet = wb.active

count = 0
for filename in os.listdir("/Users/jeffmajnaric/Google Drive/Personal/Finance/RSE 2019_2020 Paystubs"):
    if count == 25:
        break
    file_jpg = convert_from_path('/Users/jeffmajnaric/Google Drive/Personal/Finance/RSE 2019_2020 Paystubs/Paystub' + str(count) + '.pdf', dpi= 300, output_folder="/Users/jeffmajnaric/Google Drive/Personal/Finance/RSE 2019_2020 Paystubs", output_file="Paystub" + str(count), fmt='JPEG') # saved as a list
    count = count + 1

count = 0
xl_count = 2 # used for counting row number
for filename in os.listdir("/Users/jeffmajnaric/Google Drive/Personal/Finance/RSE 2019_2020 Paystubs"):
    if filename.endswith(".jpg"):
        im = Image.open('/Users/jeffmajnaric/Google Drive/Personal/Finance/RSE 2019_2020 Paystubs/Paystub' + str(count) + '0001-1.jpg')
        im = im.crop((1128, 800, 1296, 1012))
        text = pytesseract.image_to_string(im)
        print(text)
        sheet["C" + str(xl_count)] = text
        wb.save('/Users/jeffmajnaric/Google Drive/Personal/Finance/RSE_Paystubs.xlsx')
        os.remove('/Users/jeffmajnaric/Google Drive/Personal/Finance/RSE 2019_2020 Paystubs/Paystub' + str(count) + '0001-1.jpg')
        count = count + 1
        xl_count = xl_count + 1 
