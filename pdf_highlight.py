from PyPDF2 import PdfFileWriter, PdfFileReader
from PyPDF2.generic import (
    DictionaryObject,
    NumberObject,
    FloatObject,
    NameObject,
    TextStringObject,
    ArrayObject
)

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument, PDFNoOutlines
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTFigure, LTImage, LTChar
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import resolve1

import re
import xlwings as xw
import pathlib
import random
from os import listdir


# extracting coordinates from text line
def get_coordinates(lt_obj, query):
    coor = []
    result = re.finditer(query, lt_obj.get_text())
    for match in result:

        s = match.start()
        e = match.end()-1

        fl_coor = lt_obj._objs[s].bbox
        ll_coor = lt_obj._objs[e].bbox
        s_word = (fl_coor[0], fl_coor[1], ll_coor[2], ll_coor[3])
        coor.append(s_word)
    return coor


# functions
# for highlight
# x1, y1 starts in bottom left corner
def create_highlight(x1, y1, x2, y2, meta, color = [0.5, 0, 0]):
    newHighlight = DictionaryObject()

    newHighlight.update({
        NameObject("/F"): NumberObject(4),
        NameObject("/Type"): NameObject("/Annot"),
        NameObject("/Subtype"): NameObject("/Highlight"),

        NameObject("/T"): TextStringObject(meta["author"]),
        NameObject("/Contents"): TextStringObject(meta["contents"]),

        NameObject("/C"): ArrayObject([FloatObject(c) for c in color]),
        NameObject("/Rect"): ArrayObject([
            FloatObject(x1),
            FloatObject(y1),
            FloatObject(x2),
            FloatObject(y2)
        ]),
        NameObject("/QuadPoints"): ArrayObject([
            FloatObject(x1),
            FloatObject(y2),
            FloatObject(x2),
            FloatObject(y2),
            FloatObject(x1),
            FloatObject(y1),
            FloatObject(x2),
            FloatObject(y1)
        ]),
    })
    return newHighlight


# get coordinates from page
def get_page_coordinates(page, query):

    word_coor = []
    for lt_obj in page:
        if isinstance(lt_obj, LTTextLine):
            result = get_coordinates(lt_obj, query)
            for item in result:
                word_coor.append(item)
        elif isinstance(lt_obj, LTTextBox):
            for line in lt_obj:
                result = get_coordinates(line, query)
                for item in result:
                    word_coor.append(item)

    return word_coor



def anotate_pdf(file_path, sht, query_dict):

    # preparing the output file name
    path = pathlib.Path(file_path).parent
    extension = pathlib.Path(file_path).suffix
    name = pathlib.Path(file_path).name[:-len(extension)]
    result_file = str(path)+'\\'+name+'_highlighted'+extension

    #=========================================================

    # create a parser object associated with the file object
    parser = PDFParser(open(file_path, 'rb'))
    # create a PDFDocument object that stores the document structure
    doc = PDFDocument(parser)

    # Layout Analysis
    # Set parameters for analysis.
    laparams = LAParams()
    # Create a PDF page aggregator object.
    rsrcmgr = PDFResourceManager()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    # create pdf layout - this is list with layout of every page
    layout = []
    for page in PDFPage.create_pages(doc):
        interpreter.process_page(page)
        # receive the LTPage object for the page.
        layout.append(device.get_result())


    # add tooltip info not sure how to use this option in the most usefull way
    m_meta = {"author": "AK",
              "contents": "HL text1"}

    outputStream = open(result_file, "wb")
    pdfInput = PdfFileReader(open(file_path, 'rb'), strict=True)
    pdfOutput = PdfFileWriter()


    npage = pdfInput.numPages
    for pgn in range(0, npage):
        for query in query_dict:
            all_coor = []
            for page in layout:
                result = get_page_coordinates(page, query)
                all_coor.append(result)

            page_hl = pdfInput.getPage(pgn)

            for item in all_coor[pgn]:
                highlight = create_highlight(item[0], item[1], item[2], item[3], m_meta, color = query_dict[query])
                highlight_ref = pdfOutput._addObject(highlight)

                if "/Annots" in page_hl:
                    page_hl[NameObject("/Annots")].append(highlight_ref)
                else:
                    page_hl[NameObject("/Annots")] = ArrayObject([highlight_ref])


        pdfOutput.addPage(page_hl)

    # save HL to new file
    pdfOutput.write(outputStream)
    outputStream.close()
    sht.range('B2').value = f'File {name+extension} completed'



def annotate_pdfs():
    # get data from excel file
    wb = xw.Book('pdf_highlight.xlsm')
    sht = wb.sheets['Sheet1']
    file_path = sht.range('B1').value

    #  check for words
    word_range = sht.range('A5').expand('down').address
    # take the words from excel
    words = sht.range(word_range).value

    #  check for colors
    # using start and end from words
    col_range = word_range.replace('A', 'B')
    colors = []
    # check which cells from B column contain data
    c_list = col_range.split(':')
    t_range = [int(x.replace('$B$', '')) for x in c_list]
    for val in range(t_range[0], t_range[1] + 1):
        checkvalue = sht.range('B' + str(val)).color
        if checkvalue is not None:
            colors.append(checkvalue)
        else:
            print(f"No color for word: {sht.range('A' + str(val)).value}, assigning random color!")
            colors.append((random.randint(0, 255), random.randint(0, 255), random.randint(0, 255)))

    # convert colors to list float 0 to 1
    colors_list = [list(x) for x in colors]
    colors = []
    for i in colors_list:
        i = [x / 255 for x in i]
        colors.append(i)

    # create query dictionary
    query_dict = {}
    for i in range(len(words)):
        query_dict[words[i]] = colors[i]


    if '.pdf' in file_path:
        anotate_pdf(file_path, sht, query_dict)
    else:
        files = listdir(file_path)
        pdf_files = [x for x in files if '.pdf' in x]

    for i, pdf in enumerate(pdf_files):
        sht.range('B2').value = f'Annotating file {pdf} (file {i+1} of {len(pdf_files)})'
        anotate_pdf(file_path + '/' + pdf, sht, query_dict)
    sht.range('B2').value = f'Annotation complete!'


annotate_pdfs()
# to do
# fix the problem with displaced highligts ('J. Biol. Chem.-2014-Kichev-9430-9')
# better text encoding special characters as alpha beta etc

