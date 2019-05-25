import xlrd
import xlwt
from xlutils.copy import copy
import re

def get_words(filename, sheet_num):
    data_to_tr = {'nouns': [], 'adjectives' : []}
    rb = xlrd.open_workbook(filename) #'questionnaire_name.xlsx'
    sheet = rb.sheet_by_index(sheet_num)
    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        for c_el in row:
            if rownum == 0 and c_el != '':
                data_to_tr['adjectives'].append(c_el)
            elif c_el != '':
                data_to_tr['nouns'].append(c_el)
    return data_to_tr

    
def find_by_lem(adjective, corpus):  
    collocations = []
    with open(corpus,'r', encoding='utf-8') as file:
        for line in file:
            if ' '+adjective+'..' in line:
                splitted = line.split(' ')
                indexes = [splitted.index(item) for item in splitted if adjective in item]
                for index in indexes:
                    if index == len(splitted)-1:
                        pass
                    elif '..nn.' in splitted[index+1]:
                        collocations.append(splitted[index+1])
    return collocations

def prepare_file_to_write(filename):
    rb = xlrd.open_workbook(filename)
    wb = copy(rb)
    return wb


def main(questionnaire, feature_name, sheet_num, corpus):
    words = get_words(questionnaire, sheet_num)
    print(words)

    wb = prepare_file_to_write(questionnaire)
    w_sheet = wb.get_sheet(sheet_num) # the sheet to write to within the writable copy
    ws = wb.add_sheet('nouns_'+feature_name)
    other_nouns = {}

    for word in words['adjectives']:
        other_nouns[word]=[]
        i = 1
        col = words['adjectives'].index(word) + 1
        print(word)
        print('processing...')
        collocations = find_by_lem(word, corpus)
        print('done')
        for noun in collocations:
            noun_no_tag = re.sub('\.\.nn..', '', noun)
            if collocations.count(noun) >= 10:
                if noun_no_tag in words['nouns']:
                    row = words['nouns'].index(noun_no_tag) + 1
                    w_sheet.write(row, col, 'yes')
                else:
                    ws.write(i, col, noun)
                    i += 1
                    other_nouns[word].append(noun)
            elif noun_no_tag in words['nouns']:
                row = words['nouns'].index(noun_no_tag) + 1
                w_sheet.write(row, col, 'rare')
        
    wb.save('filled_questionnaire.xls')
    return "questionnaire is filled"

#example for 'sharp' questionnaire
main('ankety.xlsx', 'sharp', 0, 'gigaword.txt')
    
    

