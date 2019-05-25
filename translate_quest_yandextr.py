import requests
import json
import xlrd

def yandex_translate_single(to_translate, source_language, target_language, reverse=False):
    """
    Обёртка яндексового апи
    :param to_translate:
    :param source_language:
    :param target_language:
    :param reverse:
    :return:
    """
    if reverse:
        source_language, target_language = target_language, source_language
    language_direction = f"{source_language}-{target_language}"
    translate_params = {'key': 'trnsl.1.1.20190303T082102Z.f7e151ab791be4f8.b3803e2ac222bee403a5c1e67682be8e9a682065',
                        'text': to_translate,
                        'lang': language_direction}
    translate_request = requests.get('https://translate.yandex.net/api/v1.5/tr.json/translate', translate_params)
    translation_data = json.loads(translate_request.text)
    return translation_data['text'][0]

def get_words(filename):
    data_to_tr = {'nouns': [], 'adjectives' : []}
    rb = xlrd.open_workbook(filename) #'questionnaire_size.xlsx'
    sheet = rb.sheet_by_index(2)
    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        for c_el in row:
            if rownum == 0 and c_el != '':
                data_to_tr['adjectives'].append(c_el)
            elif c_el != '':
                data_to_tr['nouns'].append(c_el)
    return data_to_tr

def translate_some_words(words):
    translations = []
    for word in words:
        translations.append(yandex_translate_single(word, 'en', 'sv'))
    return translations

def new_questionnire(data_to_tr):
    new_questionnaire = {}
    new_questionnaire['nouns'] = translate_some_words(data_to_tr['nouns'])
    new_questionnaire['adjectives'] = translate_some_words(data_to_tr['adjectives'])
    return new_questionnaire

    
    
data = get_words('анкеты на английском.xlsx')
new = new_questionnire(data)
print(new)
