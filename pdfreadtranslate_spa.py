import os
import xlsxwriter
import re
import codecs
import msmt
import time

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO


# Function convert_pdf_to_txt was imported from another project.
# Unknown author to be credited
def convert_pdf_to_txt(path):
    """
    Reads the pdf files into text streams
    """
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(
            fp, pagenos, maxpages=maxpages, password=password,caching=caching,
            check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    return text
    fp.close()
    device.close()
    retstr.close()


def scan_keywords():
    """
    Regex-based search for relevant keywords.
    Also allows for the exclusion of phrases known to produce false hits
    """
    try:
        c = len(file_list)
        j = len(ex_list)
        output = f[i][0:len(f[i])-4]
        output = output.replace(".", "_") + ".txt"
        lower_text = output.lower()

        os.chdir(os.path.join(home,fileFolder))
        parser = PDFParser(open(f[i], 'rb'))
        document = PDFDocument(parser)

        # Skips file that are secured
        if not document.is_extractable:
            print "Not Allowed: " + lower_text
        else:
            print lower_text
            text = convert_pdf_to_txt(f[i]).decode('utf-8')
            # Stores the spanish match and the translated text in separate lists
            match = list(u'')
            en_match = list(u'')

            # Checks for phrases to be excluded
            for ex in ex_keys:
                rg = re.compile(
                    ex.rstrip(), re.IGNORECASE|re.DOTALL|re.U|re.MULTILINE)
                xm = rg.search(text.lower())

                if xm is not None:
                    break

            ex_keys.seek(0)

            # Looks for desired keywords if there are no unwannted keywords
            if xm is None:
                #Looks for the needles by regex
                for r in rgx_keys:
                    rg = re.compile(
                        r.rstrip(),re.IGNORECASE|re.DOTALL|re.U|re.MULTILINE)
                    m = rg.search(text.lower())

                    if m is not None:
                        match.append(m.group(0))
                        # Translates the matched phrase to english
                        en_text = translate_text(m.group(0))
                        en_match.append(en_text)

                rgx_keys.seek(0)

                if len(match) > 0:
                    # Writes the pdf contents to text and append the results
                    file_list.append([])
                    os.chdir(os.path.join(home, destFolder))

                    with open(output, 'w') as text_file:
                        text_file.write(text.encode('utf-8'))
                        text_file.close()
                        file_list[c].append(os.path.join(destFolder, output))
                        file_list[c].append(enumerate(match))
                        file_list[c].append(enumerate(en_match))

    # Error logging
    except Exception as e:
        ex_list.append([])
        ex_list[j].append(os.path.join(destFolder, output))
        ex_list[j].append(str(e))


# Function to translate the Spanish texts to english
def translate_text(match_phrase):
    # Gets the access token to communicate with MS Translator
    client_id = 'your-registered-id'  # replace with your own id
    client_secret = 'your-client-secret'  # replace with your own secret
    access_token = msmt.get_access_token(client_id, client_secret)

    # Cleans up the phrase
    match_phrase = ' '.join(match_phrase.split())

    # Tries for a maximum of three times
    en_result = u''
    for n in range(3):
        try:
            en_result = msmt.translate(access_token, match_phrase, 'en', 'es')
        except:
            time.sleep(10)
            # Retry translate text
            translate_text(access_token, match_phrase)

        if en_result is not None:
            en_result = re.sub('(\<.*?\>)', '', en_result)
            break

    return en_result


# Function to write all results to workbook
def write_to_file(book_name):
    # Writes the results to a spreadsheet
    y = int(os.path.splitext(book_name)[0][7])
    os.chdir(home)
    book = xlsxwriter.Workbook(book_name, {'constant_memory': True})
    sheet1 = book.add_worksheet("Sheet 1")
    sheet2 = book.add_worksheet("Sheet 2")
    sheet1.write(0, 0, "Text file")
    sheet1.write(0, 1, "ES_Match")
    sheet1.write(0, 2, "EN_Match")
    sheet2.write(0, 0, "File name")
    sheet2.write(0, 1, "Error type")

    for k, v in enumerate(file_list, start = 1):
        try:
            sheet1.write_url(k, 0, v[0])
        except:
            sheet1.write_url(k, 0, 'Failed to link file')

        es_out = u''
        try:
            for l, w in v[1]:
                es_out = es_out + str(l+1) + u') ' + w + u' '
            sheet1.write(k, 1, es_out)
        except:
            sheet1.write(k, 1, "Data Write Exception")

        en_out = u''
        try:
            for m, x in v[2]:
                en_out = en_out + str(m+1) + u') ' + x + u' '
            sheet1.write(k, 2, en_out)
        except:
            sheet1.write(k, 2, 'Server Connection Failure')

    for k, v in enumerate(ex_list, start = 1):
        try:
            sheet2.write_url(k, 0, v[0])
        except:
            sheet2.write_url(k, 0, 'Failed to link file')

        try:
            sheet2.write(k, 1, v[1])
        except:
            sheet2.write(k, 1, "Data Write Exception")

    book.close()
    del file_list[:]
    del ex_list[:]
    y += 1
    return "summary" + str(y) + ".xlsx"


home = os.getcwd()
fileFolder = 'folder-to-scan'  # replace this with the name of the target folder
destFolder = 'texts'  # destination folder for text files with positive matches
rgx_keys = codecs.open(
    'rgx_spanish_keywords2.txt', encoding='utf-8-sig', mode='r')
ex_keys = codecs.open(
    'exclude_keywords.txt', encoding='utf-8-sig', mode='r')
f = []

os.chdir(os.path.join(home,fileFolder))
for (dirpath, dirnames, filenames) in os.walk(os.path.join(home,fileFolder)):
    for file_ in filenames:
        # Only reads pdf files
        # Replace the decoder with the encoding used by the local machine
        if file_.decode('cp1252').lower().endswith(".pdf"):
            f.append(file_)

file_list = []
ex_list = []
y = 1
book_name = "summary1.xlsx"


# The entire mining process is broken down into 8 parts
# Prevents the loss of progress in the event of machine hiccups
# TODO: Compress the repeated procedure below into a single looped procedure
print "Begin Part 1"

for i in range(0, int(len(f)/8)):
    scan_keywords()
book_name = write_to_file(book_name)

print "Begin Part 2"

for i in range(int(len(f)/8), int(len(f)/4)):
    scan_keywords()
book_name = write_to_file(book_name)

print "Begin Part 3"

for i in range(int(len(f)/4), int(3*len(f)/8)):
    scan_keywords()
book_name = write_to_file(book_name)

print "Begin Part 4"

for i in range(int(3*len(f)/8), int(len(f)/2)):
    scan_keywords()
book_name = write_to_file(book_name)

print "Begin Part 5"

for i in range(int(len(f)/2), int(5*len(f)/8)):
    scan_keywords()
book_name = write_to_file(book_name)

print "Begin Part 6"

for i in range(int(5*len(f)/8), int(3*len(f)/4)):
    scan_keywords()
book_name = write_to_file(book_name)

print "Begin Part 7"

for i in range(int(3*len(f)/4), int(7*len(f)/8)):
    scan_keywords()
book_name = write_to_file(book_name)

print "Begin Part 8"

for i in range(int(7*len(f)/8), len(f)):
    scan_keywords()
book_name = write_to_file(book_name)

print "Mining complete!"
