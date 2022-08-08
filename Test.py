from textblob import TextBlob
import easyocr
import warnings

imgPath = './imgs/132.png'

reader = easyocr.Reader(['en'], gpu=True)
result = reader.readtext(imgPath)

if(int(result[0][2]*10.00) >= 5):
    match = result[0][1]
    try:
        if(int(result[1][2]*10.00) >= 4):
            match = match+' '+result[1][1]
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


print(match)

text = "tiktok"
tb = TextBlob(text)
translated = tb.translate(from_lang='en', to="ar")
print(str(translated))
