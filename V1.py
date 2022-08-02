import PyPDF2
from openpyxl import load_workbook
import tkinter as tk
import arabic_reshaper


def Extract():
    message.config(text="")
    inp = PDF_FILE.get()
    try:
        pdfFileObj = open(inp+'.pdf', 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    except:
        message.config(text="Couldn't find "+inp+'.pdf', fg='red')
        return

    inp = OLD_EXCEL_FILE.get()
    try:
        book = load_workbook(inp+'.xlsx')
    except:
        message.config(text="Couldn't find "+inp+'.xlsx', fg='red')
        return

    sheet = book.active
    print(sheet)

    name = 0
    date = 0
    number = 0
    address = 0

    count = 6
    countu = 6
    countp = 6
    countr = 6
    match = ''
    page = ""
    p = ""
    for mainpage in range(3, 52):
        p += pdfReader.getPage(mainpage).extractText() + "\n"

    for i, _ in enumerate(p):
        if p[i:i + len("بلطلا مقر")] == "بلطلا مقر":
            number = i
            match = p[i+10:i+16]
            if (match.isdigit()):
                sheet[('t' + str(count))].value = int(match)
                print(p[i+10:i+17])
                count += 1

        if(p[i:i + len("بلطلا ميدقت خيرات")] == "بلطلا ميدقت خيرات") or (p[i:i + len("بلطلا ميدقت خيرات")] == "ت خيرات  :بلطلا م"):
            date = i
            match = p[i+17:i+27]
            sheet[('u' + str(countu))].value = match
            #print(match)
            countu += 1

        if(p[i:i + len("ليجستلا بلاط مسل")] == "ليجستلا بلاط مسا"):
            lsl = i + len("ليجستلا بلاط مسل")

            #while (p[lsl:lsl+len("ةيسنجلاو ناونعلا")] != "ةيسنجلاو ناونعلا"):
            #    lsl += 1
            while (p[lsl] != '\n'):
                lsl += 1
                if(p[lsl] == ':'):
                    break

            match = p[i+17:lsl]
            if(match.upper() == match.lower()):
                match = match[::-1]
            sheet[('p' + str(countp))].value = match
            print(match)
            countp += 1

    '''    if(p[i:i + len("ةيسنجلاو ناونعلا")] == "ةيسنجلاو ناونعلا"):
            lsl = i + len("ةيسنجلاو ناونعلا")

            while (p[lsl:lsl+len("الفئة")] != "ةئفلا"):
                lsl += 1

            match = p[i+16:lsl-3]
            for letter in match:
                if letter == '\n':
                    letter = ' '
            match = match[::-1]
            sheet[('r' + str(countr))].value = match
            print(match)
            countr += 1'''
    inp = NEW_EXCEL_FILE.get()
    if(inp == ''):
        inp = 'NewFile.xlsx'
    try:
        book.save(inp)
    except:
        message.config(text='Make sure the output file is closed', fg='red')
        return
    pdfFileObj.close()
    message.config(text='Done Extracting', fg='green')


# Top level window
frame = tk.Tk()
frame.title("BOT")
frame.geometry('400x200')


PDF_LABEL = tk.Label(frame, text="PDF file name ")
PDF_LABEL.place(x=0, y=0)

PDF_FILE = tk.Entry(frame)
PDF_FILE.insert(-1, '1')

PDF_FILE.place(x=150, y=0)


OLD_EXCEL_LABEL = tk.Label(frame, text="Template file name ")
OLD_EXCEL_LABEL.place(x=0, y=30)

OLD_EXCEL_FILE = tk.Entry(frame)
OLD_EXCEL_FILE.insert(-1, 'Templates')
OLD_EXCEL_FILE.place(x=150, y=30)

NEW_EXCEL_LABEL = tk.Label(frame, text="Output file name ")
NEW_EXCEL_LABEL.place(x=0, y=60)

NEW_EXCEL_FILE = tk.Entry(frame)
NEW_EXCEL_FILE.insert(-1, 'output.xlsx')
NEW_EXCEL_FILE.place(x=150, y=60)

# Button Creation
ExtractButton = tk.Button(frame,
                          text="Extract",
                          relief='groove',
                          cursor='hand2',
                          command=Extract)
ExtractButton.place(x=150, y=90)
inp = tk.StringVar()

# Label Creation
message = tk.Label(frame, text="")
message.place(x=130, y=120)


pdfFileObj = open('1.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
pageObj = pdfReader.getPage(4)
p = pageObj.extractText()

for mainpage in range(3, 52):
    p += pdfReader.getPage(mainpage).extractText() + "\n"

'''for i, _ in enumerate(p):
    if(p[i:i + len("ليجستلا بلاط مسل")] == "ليجستلا بلاط مسا"):
        lsl = i + len("ليجستلا بلاط مسل")

        while (p[lsl:lsl+len("ةيسنجلاو ناونعلا")] != "ةيسنجلاو ناونعلا"):
            lsl += 1

        match = p[i+17:lsl-3]
        match = match[::-1]
#if (match.isdigit()):
#print(p[11647:12618])
match = p[11647:12000]
print(match)'''


frame.mainloop()
