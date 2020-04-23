# PaystubToExcel

This script uses OCR to extract the net pay from each paystub and exports it all to an existing Excel spreadsheet.

I start by iterating through the directory where all the paystub files in PDF format are located. Since pytesseract only works for picture file formats, I converted them all to PNG files to be used by the OCR by using the pdf2image library. After converting all the files, I crop a specific area using the PIL library in the PNG file to be analyzed by pytesseract to extract the text as a string. Since the net pay value in each paystub is in the same position approximately for all files, I am able to do this with no issues. I then store each value in the Excel sheet and delete the PNG file after it is used.

One alternative that I was initially attempting was to extract all of the text from the PNG files without cropping. I would then have to parse the text to find the specific net pay value. I opted for cropping and then applying the OCR due to its simplicity.

The code can still be optimized as I am running two separate for loops to turn all the PDF files to PNG format and to iterate through each PNG file and extract the text. As a result of trying this way, the code was not properly deleting the files when going through. Another aspect for optimization is simply avoiding the step of converting the PDF files to PNG but I would have to find a different library to use this format.
