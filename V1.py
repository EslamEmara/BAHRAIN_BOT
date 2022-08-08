from datetime import date
import PyPDF2
import openpyxl
import tkinter as tk
from tkinter import filedialog
from tkinter.ttk import Progressbar
from openpyxl.styles import Alignment
from openpyxl.styles import Color, PatternFill
from openpyxl.styles import colors
import os
import PIL
import time
import pdfminer
from pdfminer.image import ImageWriter
from pdfminer.high_level import extract_pages
import pytesseract
from textblob import TextBlob
import easyocr
import warnings
warnings.filterwarnings("ignore", category=UserWarning)

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

location = './imgs'
pdfLocation = './'

try:
    dir_list = os.listdir(location)
except:
    os.mkdir(location)
    dir_list = os.listdir(location)


logo = PIL.Image.open('./Dependencies/mainLogo.png')
hyp = PIL.Image.open('./Dependencies/Hyperlogo.png')


today = date.today()

# dd/mm/YY
d1 = today.strftime("%d/%m/%Y")


def openPDFFile(inp):
    global miner_pages
    try:
        pdfFileObj = open(inp, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        if From_entry.get() == '':
            From_entry.insert(-1, 1)
        if To_entry.get() == '':
            To_entry.insert(-1, pdfReader.numPages)

        return pdfReader, pdfFileObj
    except:
        return -1


def get_image(layout_object):
    if isinstance(layout_object, pdfminer.layout.LTImage):
        return layout_object
    if isinstance(layout_object, pdfminer.layout.LTContainer):
        for child in layout_object:
            return get_image(child)
    else:
        return None


def openExcelFile(inp):
    try:
        book = openpyxl.load_workbook(OLD_EXCEL_FILE.get())
        return book
    except:
        return -1


def ExtractAllPages(inp, pdfReader):
    page = ""

    try:
        first_page = int(From_entry.get())-1
        last_page = int(To_entry.get())+1
    except:
        return -1

    if ((first_page < 0) or (last_page > pdfReader.numPages+1)):

        return -1

    for mainpage in range((first_page), last_page):
        page += (pdfReader.getPage(mainpage).extractText() + "\n")

    progress_bar.step(30)
    frame.update()
    miner_pages = list(extract_pages(inp))

    directory = r'./imgs/'
    counter = 2
    logo = PIL.Image.open('./Dependencies/mainLogo.png')
    hyp = PIL.Image.open('./Dependencies/Hyperlogo.png')

    for i in range(first_page, last_page):
        min_page = miner_pages[i]
        images = list(filter(bool, map(get_image, min_page)))
        iw = ImageWriter('./imgs/')

        for image in images:
            image.name = str(counter)
            iw.export_image(image)

            try:
                im = PIL.Image.open(directory+image.name+'.jpg')
                name = directory+image.name+'.jpg'
            except:
                im = PIL.Image.open(directory+image.name+'.bmp')
                name = directory+image.name+'.bmp'

            rgb_im = im.convert('RGB')
            pngName = directory+image.name+'.png'
            rgb_im.save(pngName)

            os.remove(name)

            img = PIL.Image.open(pngName)
            if list(logo.getdata()) == list(img.getdata()) or (list(hyp.getdata()) == list(img.getdata())):
                os.remove(pngName)
                continue

            counter += 1
            frame.update()

    progress_bar.step(40)
    frame.update()

    return page


def ReshapeArabicText(p):
    p = p[::-1]
    list1 = list(p)
    done = 0
    done2 = 0
    for c, _ in enumerate(list1):
        if(list1[c].isdigit()):
            if(done == 0):
                counter = 0
                while True:
                    if((list1[c+counter]).isdigit()):
                        counter += 1
                    else:
                        list1[c:c+counter] = list1[c:c+counter][::-1]
                        done = 1
                        done2 = 0
                        break

        elif(str(list1[c]).upper() != str(list1[c]).lower()):
            if(done2 == 0):
                counter = 0
                while True:
                    if(str(list1[c+counter]).upper() != str(list1[c+counter]).lower()):
                        counter += 1
                    else:
                        list1[c:c+counter] = list1[c:c+counter][::-1]
                        done2 = 1
                        done = 0
                        break
        else:
            done2 = 0
            done = 0
    p = ''.join(list1)

    return p


def splitText(p):
    LinesList = p.split('\n')
    ColonsList = []
    for i in LinesList:
        ColonsList.append(i.split(':'))
    ColonsList.reverse()

    return ColonsList


def ParsingReqText(a, sheet):
    newLines = 0
    rowTcount = 2
    rowUcount = 2
    rowPcount = 2
    rowRcount = 2
    rowVcount = 2
    rowAJcount = 2
    rowAEcount = 2

    rowAcount = 2
    match = ''
    PublicationNumber = ''
    PubDate = ''
    figures = 0
    f1 = 0
    f2 = 0
    f3 = 0
    f4 = 0
    f5 = 0
    f6 = 0
    f7 = 0
    #print(a)
    numberPerPage = 0
    addressDone = 1
    pageNumber = -1
    dir_list = os.listdir(location)
    users_per_page = []
    images_per_page = []
    for c, j in enumerate(a):
        for c2, k in enumerate(j):
            lstr = ''.join(a[c])
            #print(lstr)
            if "رقم الطلب" in k:
                match = k.replace("رقم الطلب", "")
                try:
                    sheet[('s' + str(rowTcount))].value = int(match)
                except:
                    message.config(text="")
                    message.config(
                        text="Choose proper pages for high accuracy", fg='orange')
                    frame.update()
                    time.sleep(2)
                    break
                sheet[('s' + str(rowTcount))
                      ].alignment = Alignment(horizontal='right')
                sheet.row_dimensions[rowTcount].height = 150

                imgPath = './imgs/'+str(rowTcount)+'.png'

                reader = easyocr.Reader(['en'], gpu=True)
                result = reader.readtext(imgPath)
                try:
                    if(int(result[0][2]*10.00) >= 5):
                        match = result[0][1]
                        try:
                            if(int(result[1][2]*10.00) >= 4):
                                match = match+' '+result[1][1]
                        except:
                            pass
                        sheet[('a' + str(rowTcount))].value = match
                        tb = TextBlob(match)
                        try:
                            translated = tb.translate(from_lang='eng', to="ar")
                            sheet[('b' + str(rowTcount))
                                  ].value = str(translated)
                        except:
                            pass
                    else:
                        reader = easyocr.Reader(['ar'], gpu=False)
                        result = reader.readtext(imgPath)
                        if(int(result[0][2]*10.00) >= 5):
                            match = result[0][1]
                            try:
                                if(int(result[1][2]*10.00) >= 4):
                                    match = match+' '+result[1][1]
                            except:
                                pass
                        else:
                            match = "none"

                    sheet[('b' + str(rowTcount))].value = match
                    tb = TextBlob(match)
                    try:
                        translated = tb.translate(from_lang='ar', to="eng")
                        sheet[('a' + str(rowTcount))].value = str(translated)
                    except:
                        pass

                except:
                    pass

                try:
                    img = openpyxl.drawing.image.Image(
                        './imgs/'+str(rowTcount)+'.png')
                    img.anchor = 'H' + str(rowTcount)
                    sheet.add_image(img)
                except:
                    pass
                rowTcount += 1
                numberPerPage += 1
                f1 = 1

            if "الفئة" in k:
                match = k.replace("الفئة", "")
                match = match.replace(" ", "")
                if (match.isdigit()):
                    sheet[('u' + str(rowVcount))].value = int(match)
                    sheet[('u' + str(rowVcount))
                          ].alignment = Alignment(horizontal='right')
                    rowVcount += 1
                    f2 = 1

            if "تاريخ تقديم الطلب" in k or "يم الطلب" in k:
                if("تاريخ تقديم الطلب" in k):
                    match = k.replace("تاريخ تقديم الطلب", "")
                else:
                    match = k.replace("يم الطلب", "")
                sheet[('t' + str(rowUcount))].value = match
                sheet[('t' + str(rowUcount))
                      ].alignment = Alignment(horizontal='right')
                rowUcount += 1
                f3 = 1

            if "اسم طالب التسجيل" in k:
                match = k.replace("اسم طالب التسجيل", "")
                match = match.replace("اإل", "الا")
                match = match.replace("األ", "الا")
                match = match.replace("اال", "الا")

                match = match.replace("اإل", "الا")
                match = match.replace("األ", "الا")
                match = match.replace("اال", "الا")
                sheet[('o' + str(rowPcount))].value = match
                sheet[('o' + str(rowPcount))
                      ].alignment = Alignment(horizontal='right')
                rowPcount += 1
                f4 = 1

            if "والجنسية" in k or "العنوان والجنسية" in k:
                if ("العنوان والجنسية" in k):
                    match = ''.join(a[c]).replace("العنوان والجنسية", "")
                else:
                    match = ''.join(a[c]).replace("والجنسية", "")

                match = match + ''.join(a[c+1])
                match = match.replace("اإل", "الا")
                match = match.replace("األ", "الا")
                match = match.replace("اال", "الا")
                if("عنوان" in (''.join(a[c+2])) or 'الفئة' in (''.join(a[c+2])) or 'اسم الوكيل' in (''.join(a[c+2])) or "اااااااااااااااااااااااااااااااااااااااااااااااااااااااااااا" in (''.join(a[c+2])) or 'ا لعدد' in (''.join(a[c+2]))):
                    pass
                else:
                    match = match + ''.join(a[c+2])

                if ("العنوان" in match):
                    match = match.replace("العنوان", "")
                if ("اسم طالب التسجيل" in match):
                    match = match.replace("اسم طالب التسجيل", "")
                if ("تيك توك إل تي دي." in match):
                    match = match.replace("تيك توك إل تي دي.", "")

                sheet[('q' + str(rowRcount))].value = match
                sheet[('q' + str(rowRcount))
                      ].alignment = Alignment(horizontal='right')
                tstring = ''

                if (',' in match):
                    tstring = match.split(',')[-1]
                    if(len(tstring) >= 3):
                        match = tstring
                    else:
                        match = match.split(',')[-2]

                elif("،" in match):
                    tstring = match.split("،")[-1]
                    if(len(tstring) >= 3):
                        match = tstring
                    else:
                        match = match.split('،')[-2]
                else:
                    match = ""

                try:
                    match = ''.join([i for i in match if not i.isdigit()])
                except:
                    pass

                sheet[('r' + str(rowRcount))].value = match
                sheet[('r' + str(rowRcount))
                      ].alignment = Alignment(horizontal='right')

                #tr = translator.translate(match, src='ar', dest='en')
                #print(tr.text)
                rowRcount += 1
                f5 = 1

            if "البضائع/الخدمات" in k:
                match = ''.join(a[c]).replace("البضائع/الخدمات", "")
                match = match.replace('الفئة', "")
                newLines = 1
                while True:
                    try:
                        if("عنوان" in (''.join(a[c+newLines])) or 'الفئة' in (''.join(a[c+newLines])) or 'اسم الوكيل' in (''.join(a[c+newLines])) or "اااااااااااااااااااااااااااااااااااااااااااااااااااااااااااا" in (''.join(a[c+newLines])) or 'ا لعدد' in (''.join(a[c+newLines]))):
                            break
                        else:
                            match = match + ''.join(a[c+newLines])
                            newLines += 1
                    except:
                        break

                match = match.replace("اإل", "الا")
                match = match.replace("األ", "الا")
                match = match.replace("اال", "الا")

                sheet[('ai' + str(rowAJcount))].value = match
                sheet[('ai' + str(rowAJcount))
                      ].alignment = Alignment(horizontal='right')
                rowAJcount += 1
                f6 = 1

            if (("سم الو" in lstr) or ("الوكيل   اسم" in lstr)):

                #print(k)
                if "اسم الوكيل" in ''.join(a[c]):
                    match = ''.join(a[c]).replace("اسم الوكيل", "")
                elif "كيل  اسم الو" in ''.join(a[c]):
                    match = ''.join(a[c]).replace("كيل  اسم الو", "")
                elif "الوكيل   اسم" in ''.join(a[c]):
                    match = ''.join(a[c]).replace("الوكيل   اسم", "")
                elif "سم الوكيل" in lstr:
                    match = lstr.replace("سم الوكيل", "")
                else:
                    match = k.replace("سم الوكيل", "")

                if "عنوان" in ''.join(a[c]):
                    ssss = (''.join(a[c]).split("عنوان الوكيل"))
                    try:
                        match = ssss[1]
                    except:
                        match = ""
                    match = match.replace("اسم الوكيل", "")

                    addressDone = 0
                match = match.replace("اإل", "الا")
                match = match.replace("األ", "الا")
                match = match.replace("اال", "الا")

                sheet[('AD' + str(rowAEcount))].value = match
                sheet[('AD' + str(rowAEcount))
                      ].alignment = Alignment(horizontal='right')

                rowAEcount += 1
                f7 = 1

                if addressDone == 1:
                    break

            if ("عنوان الوكيل" in lstr or "الوكيل عنوان" in lstr or "ان الوكيل عنو" in lstr or "كيل عنوان الو" in lstr):
                #if ("عنوان الوكيل" in lstr):
                if("عنوان الوكيل" in lstr):
                    match = lstr.replace("عنوان الوكيل", "")
                elif("الوكيل عنوان" in lstr):
                    match = lstr.replace("الوكيل عنوان", "")
                elif("ان الوكيل عنو" in lstr):
                    match = lstr.replace("ان الوكيل عنو", "")
                elif("كيل عنوان الو" in lstr):
                    match = lstr.replace("كيل عنوان الو", "")

                if "اسم الوكيل" in lstr:
                    match = (''.join(a[c]).split("عنوان الوكيل"))[0]
                    match = match.replace("اسم الوكيل", "")
                    addressDone = 1

                newLines = 1
                while True:
                    try:
                        if("اااااااااااااااااااااااااااااااااااااااااااااااااااااااااااا" in (''.join(a[c+newLines])) or "ا لعدد" in (''.join(a[c+newLines])) or "اسم الوكيل" in (''.join(a[c+newLines]))):
                            break
                        else:
                            match = match + ''.join(a[c+newLines])
                            newLines += 1
                    except:
                        break
                match = match.replace("اإل", "الا")
                match = match.replace("األ", "الا")
                match = match.replace("اال", "الا")
                sheet[('AF'+str(rowAcount))].value = match
                sheet[('AF'+str(rowAcount))
                      ].alignment = Alignment(horizontal='right')

                rowAcount += 1
                f8 = 1
                addressDone = 1
                break

            if ("ا لعدد" in lstr):
                match = lstr.replace("ا لعدد", "")
                PublicationNumber = match.split('–')[0]
                PubDate = match.split('–')[1]
                break

            sheet[('AA'+str(rowTcount))].value = d1
            sheet[('AL'+str(rowTcount))].value = PublicationNumber
            sheet[('AJ'+str(rowTcount))].value = PubDate
            sheet[('J'+str(rowTcount))].value = 'Bahrain'
            sheet[('L'+str(rowTcount))].value = 'Published'

            if("ااااااااااااااااااااااااااااااااااااااااااااااااااا" in k):
                if(f1 == 1):
                    f1 = 0
                else:
                    rowTcount += 1
                if(f2 == 1):
                    f2 = 0
                else:
                    rowVcount += 1
                if(f3 == 1):
                    f3 = 0
                else:
                    rowUcount += 1
                if(f4 == 1):
                    f4 = 0
                else:
                    rowPcount += 1
                if(f5 == 1):
                    f5 = 0
                else:
                    rowRcount += 1
                if(f6 == 1):
                    f6 = 0
                else:
                    rowAJcount += 1
                if(f7 == 1):
                    f7 = 0
                else:
                    rowAEcount += 1
                if(f8 == 1):
                    f8 = 0
                else:
                    rowAcount += 1

    sheet[('AA'+str(rowTcount))].value = ""
    sheet[('AL'+str(rowTcount))].value = ""
    sheet[('AJ'+str(rowTcount))].value = ""
    sheet[('J'+str(rowTcount))].value = ''
    sheet[('L'+str(rowTcount))].value = ''


def SaveExcelFile(book):
    inp = NEW_EXCEL_FILE.get()
    Alignment(horizontal='right')
    if(inp == ''):
        inp = 'NewFile.xlsx'
    try:
        book.save(inp)
    except:
        return -1

    return inp


def browseFiles():
    filename = filedialog.askopenfilename(initialdir="./",
                                          title="Select a PDF file",
                                          filetypes=(("PDF files",
                                                      "*.pdf*"),
                                                     ("All files",
                                                         "*.*")))

    # Change label contents
#    label_file_explorer.configure(text="File Opened: "+filename)
    PDF_FILE.delete(0, 'end')
    PDF_FILE.insert(-1, filename)
    openPDFFile(filename)
    pdfLocation = filename
    print(filename)


def browseFilesTemp():
    filename = filedialog.askopenfilename(initialdir="./",
                                          title="Select a Template",
                                          filetypes=(("XlSX files",
                                                      "*.xlsx*"),
                                                     ("All files",
                                                         "*.*")))

    # Change label contents
#    label_file_explorer.configure(text="File Opened: "+filename)
    OLD_EXCEL_FILE.delete(0, 'end')
    OLD_EXCEL_FILE.insert(-1, filename)
    print(filename)


def Extract_button():
    message.config(text="")
    inp = PDF_FILE.get()
    if(openPDFFile(inp) == -1):
        message.config(text="Couldn't find "+inp+'.pdf', fg='red')
        return
    else:
        pdfReader, pdfFileObj = openPDFFile(inp)

    inp = OLD_EXCEL_FILE.get()
    if(openExcelFile(inp) == -1):
        message.config(text="Couldn't find "+inp+'.xlsx', fg='red')
        return
    else:
        book = openExcelFile(inp)
        sheet = book.active

    progress_bar.place(x=100, y=170)
    frame.update()

    pages = ExtractAllPages(PDF_FILE.get(), pdfReader)
    if(pages == -1):
        message.config(text="Enter valid boundaries", fg='red')
        return
    frame.update()

    pages = ReshapeArabicText(pages)
    textList = splitText(pages)
    ParsingReqText(textList, sheet)
    sheet.column_dimensions['H'].width = 40
    progress_bar.step(20)
    frame.update()

    inp = NEW_EXCEL_FILE.get()
    if (SaveExcelFile(book) == -1):
        message.config(text='Make sure the output file is closed', fg='red')
        frame.update()
        old_time = time.time()
        while(True):
            if (SaveExcelFile(book) != -1):
                message.config(text='Done Extracting', fg='green')
                inp = SaveExcelFile(book)
                break
            if(time.time()-old_time > 20):
                frame.quit()
                break
            message.config(text='Make sure the output file is closed '
                           + str(20-(int(time.time()-old_time))), fg='red')
            frame.update()
            time.sleep(1)
    else:
        message.config(text='Done Extracting', fg='green')
        inp = SaveExcelFile(book)

    progress_bar.step(10)
    frame.update()

    pdfFileObj.close()
    os.startfile(inp)
    dir_list = os.listdir(location)
    #for files in dir_list:
    #    os.remove(location+'/'+files)


# Top level window
frame = tk.Tk()
frame.title("BOT")
frame.geometry('400x200')

#pdf to be read
PDF_LABEL = tk.Label(frame, text="PDF file path")
PDF_LABEL.place(x=0, y=5)

PDF_FILE = tk.Entry(frame)
PDF_FILE.insert(-1, 'C:/Users/20109/OneDrive/Desktop/BOT_Project/1.pdf')
PDF_FILE.place(x=150, y=5)

#first read page
From = tk.Label(frame, text="From ")
From.place(x=150, y=30)

From_entry = tk.Entry(frame, width=3)
From_entry.place(x=185, y=30)
#From_entry.insert(-1, '3')

#last read page
To = tk.Label(frame, text="To ")
To.place(x=210, y=30)

To_entry = tk.Entry(frame, width=3)
To_entry.place(x=230, y=30)
#To_entry.insert(-1, '53')

#template file name
OLD_EXCEL_LABEL = tk.Label(frame, text="Template file path ")
OLD_EXCEL_LABEL.place(x=0, y=60)

OLD_EXCEL_FILE = tk.Entry(frame)
OLD_EXCEL_FILE.insert(-1,
                      'C:/Users/20109/OneDrive/Desktop/BOT_Project/Dependencies/Template.xlsx')
OLD_EXCEL_FILE.place(x=150, y=60)
button_explore_temp = tk.Button(frame,
                                text="Browse",
                                command=browseFilesTemp,
                                cursor='hand2',
                                relief='ridge',
                                height=1)

button_explore_temp.place(x=280, y=57)
#Saving file name
NEW_EXCEL_LABEL = tk.Label(frame, text="Output file name ")
NEW_EXCEL_LABEL.place(x=0, y=90)

NEW_EXCEL_FILE = tk.Entry(frame)
NEW_EXCEL_FILE.insert(-1, 'output.xlsx')
NEW_EXCEL_FILE.place(x=150, y=90)

# Button Creation
ExtractButton = tk.Button(frame,
                          text="Extract",
                          relief='groove',
                          cursor='hand2',
                          command=Extract_button)
ExtractButton.place(x=150, y=120)

button_explore = tk.Button(frame,
                           text="Browse",
                           command=browseFiles,
                           cursor='hand2',
                           relief='ridge',
                           height=1)

button_explore.place(x=280, y=1)
inp = tk.StringVar()

# Label Creation
message = tk.Label(frame, text="")
message.place(x=130, y=150)

progress_bar = Progressbar(frame, orient='horizontal', length=200,
                           mode="determinate", takefocus=True, maximum=100)


frame.mainloop()
