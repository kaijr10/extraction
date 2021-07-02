"""
Instruction to run the script:
1. Required Python 3
2. Required pip3:
    apt install python3-pip
3. Install pdfminer for reading PDF file
    pip3 install pdfminer.six
4. Install pandas for write xml file
    pip3 install openpyxl xlsxwriter xlrd
    pip3 install pandas
5. Run the script:
    python3 extract.py <folder containing pdf files to extract> <output_excel_file>
    eg.
    python3 extract.py "folder contain pdf files" test.xlsx
"""


from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
from openpyxl import load_workbook

import pandas as pd
import re
import glob
import sys
import os
import argparse


BESTELLING = "Bestelling:"
DATUM_BESTELLING = "Datum bestelling:"
REF = "ref:"
TOTAL = "Totaalbedrag Exc BTW:"
EURO_SIGN = "€"

row = 0

def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'

    """
    laparams=LAParams(all_texts=True, detect_vertical=True, 
                      line_overlap=0.5, char_margin=1000.0, #set char_margin to a large number
                      line_margin=0.5, word_margin=5,
                      boxes_flow=1)
    """
    laparams=LAParams(char_margin=10000.0, line_margin=10)
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text

def process_text(text):
    # remove "\n"
    #text = re.sub(r"\n+", "\n", text)
    # remove empty lines
    text = "\n".join([line for line in text.split("\n") if line.strip() != ''])
    # spaces
    #text = re.sub(r"\s+", "", text)
    # split by "Bestselling:"
    items = text.split("Bestelling:")
    # list result
    result = []
    track_number = 1
    flag_stop = False
    for item in items[1:]:
        # remove orphan item
        #if "Discount Remake" in item:
        #    continue
        item = item.strip()
        # bring back "Bestelling:" to item
        item = "Bestelling:" + item
        #print(item[1])
        line_number = 1
        each_items = []
        for line_item in item.split("\n"):
            if flag_stop == True:
                continue
            if "Levering" in line_item:
                flag_stop = True
                continue
            # line 1 for Bestelling, Datum bestelling, ref
            # line 2 for Totallbedrag Exc BTW
            # remanin lines for detil
            # merge mutiple spaces to one space
            line_item = re.sub(r"\s+", " ", line_item).strip()
            #print(line_item)
            each_item = []
            if line_number == 1:
                bestelling_value = split_to_get_value(line_item, BESTELLING, DATUM_BESTELLING)
                datum_value = split_to_get_value(line_item, DATUM_BESTELLING, REF)
                ref_value = line_item.split(REF)[-1].strip()
                each_items.append(bestelling_value)
                each_items.append(datum_value)
                each_items.append(ref_value)
            elif line_number == 2:
                total_value = line_item.split(TOTAL)[1].replace(":", "").strip()
                each_items.append(total_value)
            else:
                each_time_child = []
                if "[" in line_item:
                # find index of charactar "["
                    line_item = re.split(r'%', line_item)
                    #line_item_value = " ".join(line_item[0].split(" ")[:-1])
                    line_item_value = "{line_item_0} [{value_mm}".format(
                        line_item_0=" ".join(line_item[0].split(" ")[:-1]),
                        value_mm=line_item[1].split("[")[-1]
                    )
                    #line_item = re.split(r'\d{1}mm', line_item)
                    #line_item_value = "{line_item_0}mm [{value_mm}".format(line_item_0=line_item[0],
                    #                                                    value_mm=line_item[1].split("[")[-1])
                    group_values = line_item[1].split("[")[0].strip().replace(EURO_SIGN, "").split(" ")
                    each_item_child = [value for value in group_values if value != '']
                    each_item_child.insert(0, line_item[0].split(" ")[-1].strip() + "%")
                    each_item_child.insert(0, line_item_value)
                    each_item.append(each_item_child)
                else:
                    line_item = line_item.replace(EURO_SIGN, "")
                    line_item = re.sub(r'\s+', ' ', line_item).strip()
                    line_item = line_item.split(" ")
                    each_item_child = [" ".join([line_item[0], line_item[1], line_item[2]])] + [line_item[-3], line_item[-2], line_item[-1]]
                    each_item.append(each_item_child)
                each_items.append(each_item)
            line_number = line_number + 1
        result.append(each_items)
        track_number += 1
    #print(result)
    return result

def split_to_get_value(text, keyword_start, keyword_end):
    value = text[text.find(keyword_start) + len(keyword_start):text.rfind(keyword_end)]
    return value.replace(",", "").strip()

def save_to_excel_file(result, output_file, truncate_sheet=False, sheet_name='Sheet1', startrow=None):
    global row
    # handle result
    data = {}
    bestelling = []
    datum = []
    ref = []
    totaal_btc = []
    count = []
    line_item = []
    dimension1 = []
    dimension2 = []
    btw = []
    prij = []
    totaal = []
    for item in result:
        for child_item in item[-1]:
            bestelling.append(item[0])
            datum.append(item[1])
            ref.append(item[2])
            totaal_btc.append(item[3])
            group_values = child_item[0].split(" x ")
            if "[" in child_item[0]:
                count.append(group_values[0].strip())
                line_item.append(group_values[1].split("[")[0].strip())
                dimension1.append(group_values[1].split("[")[1].strip())
                dimension2.append(group_values[2].replace("mm]", "").strip())
            else:
                line_item.append(group_values[1].strip())
                dimension1.append(0)
                dimension2.append(0)
            btw.append(child_item[1])
            prij.append(child_item[2])
            totaal.append(child_item[3])
    data['Bestelling'] = bestelling
    data['Datum bestelling'] = datum
    data['ref'] = ref
    data['totaalbedrag exc BTW'] = totaal_btc
    data['line item'] = line_item
    data['dimension 1 (mm)'] = dimension1
    data['dimension 2 (mm)'] = dimension2
    data['btw'] = btw
    data['prijs (ex)'] = prij
    data['totaal (ex)'] = totaal

    # save to output file
    df = pd.DataFrame(data)
    if not os.path.isfile(output_file):
        df.to_excel(output_file, index=None, sheet_name='Sheet1', startrow=0)
    else:
        # append it
        writer = pd.ExcelWriter(output_file, engine='openpyxl', mode='a')

        # try to open an existing workbook
        writer.book = load_workbook(output_file)
        
        # get the last row in the existing Excel sheet
        # if it was not specified explicitly
        if startrow is None and sheet_name in writer.book.sheetnames:
            startrow = writer.book[sheet_name].max_row

        # truncate sheet
        if truncate_sheet and sheet_name in writer.book.sheetnames:
            # index of [sheet_name] sheet
            idx = writer.book.sheetnames.index(sheet_name)
            # remove [sheet_name]
            writer.book.remove(writer.book.worksheets[idx])
            # create an empty sheet [sheet_name] using old index
            writer.book.create_sheet(sheet_name, idx)
        
        # copy existing sheets
        writer.sheets = {ws.title:ws for ws in writer.book.worksheets}

        if startrow is None:
            startrow = 0

        # write out the new sheet
        df.to_excel(writer, sheet_name, startrow=startrow, index=False, header=None)

        # save the workbook
        writer.save()


def parser_agrs():
    # initiate parser
    parser = argparse.ArgumentParser(description="The script tool to extract pdf files into excel file report")

    # add arguments
    parser.add_argument('-i', '--input', type=str, required=True, help="The folder contains pdf files to be extracted")
    parser.add_argument('-o', '--output', type=str, required=True, help="The output file xlsx file")
    parser.add_argument('-t', '--type', type=str, required=True, help="The type format is in : {\"invoice\", \"anwis\", \"toppoint\"}")
    
    args = parser.parse_args()
    
    return args


def main(args):
    pdf_folder = args.input
    output_file = args.output
    if not os.path.isdir(pdf_folder):
        print("ERROR - foder {pdf_folder} not exist or incorrect!".format(pdf_folder=pdf_folder))
    
    for pdf_file in glob.glob(pdf_folder + "/*.pdf"):
        try:
            # read pdf file
            text = convert_pdf_to_txt(pdf_file)
            # process the data
            result = process_text(text)
            save_to_excel_file(result, output_file)
            print("{pdf_file} completed without any issues".format(pdf_file=pdf_file))
        except Exception:
            print("{pdf_file} - have issued".format(pdf_file=pdf_file))
            pass


if __name__ == "__main__":
    main(parser_agrs())