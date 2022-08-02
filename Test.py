'''
import fitz
from PIL import Image
import io
#opeing the file
#file_path = input("Enter the PDF file path")
pdf_file = fitz.open('1.pdf')
count = 0
#Reading the location where to save the file
#location = input("Enter the location to save: ")
location = './imgs'
#finding number of pages in the pdf
number_of_pages = len(pdf_file)
logo = Image.open('mainLogo.png')
hyp = Image.open('Hyperlogo.png')
#iterating through each page in the pdf
for current_page_index in range(3, 53):
  #iterating through each image in every page of PDF
  for img_index, img in enumerate(pdf_file.get_page_images(current_page_index)):
      xref = img[0]
      image = fitz.Pixmap(pdf_file, xref)
      #if it is a is GRAY or RGB image
      if image.n >= 5:
          image = fitz.Pixmap(fitz.csRGB, image)
      img = Image.open(io.BytesIO(image.tobytes()))

      if list(logo.getdata()) != list(img.getdata()) and (list(hyp.getdata()) != list(img.getdata())):
         image.save("{}/image{}-{}.png".format(location,
                    current_page_index, img_index))
         count += 1


print(count)

'''
'''
import PyPDF2
from PIL import Image
import os

location = './imgs'
logo = Image.open('mainLogo.jpg')
hyp = Image.open('Hyperlogo.png')
input1 = PyPDF2.PdfFileReader(open("1.pdf", "rb"))
count = 0
for p in range(3, 53):
    page0 = input1.getPage(p)
    xObject = page0['/Resources']['/XObject'].get_object()

    for obj in xObject:
        if xObject[obj]['/Subtype'] == '/Image':
            size = (xObject[obj]['/Width'], xObject[obj]['/Height'])
            data = xObject[obj].getData()
            if xObject[obj]['/ColorSpace'] == '/DeviceRGB':
                mode = "RGB"
            else:
                mode = "P"

            if xObject[obj]['/Filter'] == '/FlateDecode':
                img = Image.frombytes(mode, size, data)
                if list(logo.getdata()) != list(img.getdata()) and (list(hyp.getdata()) != list(img.getdata())):
                    img.save(location+'/' + str(count) + ".png")
                    count += 1

            elif xObject[obj]['/Filter'] == '/DCTDecode':
                img = open(location+'/'+str(count) + ".jpg", "wb")
                img.write(data)
                img.close()

                img = Image.open(location+'/'+str(count) + ".jpg")
                if list(logo.getdata()) != list(img.getdata()) and (list(hyp.getdata()) != list(img.getdata())):
                    count += 1
                else:
                    os.remove(location+'/'+str(count) + ".jpg")

'''
from pikepdf import Pdf, PdfImage, Name
from PIL import Image
example = Pdf.open('1.pdf')
location = './imgs'
#logo = Image.open('mainLogo.jpg')
#hyp = Image.open('Hyperlogo.png')
for p in range(3, 53):
    page1 = example.pages[p]
    for c, i in enumerate(list(page1.images.keys())):
        print(i)
        rawimage = page1.images[i]
        pdfimage = PdfImage(rawimage)
        #pdfimage.extract_to(fileprefix=i)
        img = pdfimage.as_pil_image()
        img.save(location+'/' + str(p) + str(c) + ".png")
