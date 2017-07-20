# -- coding: utf-8 --
from op import *
from selenium import webdriver
from ws import *
import time, bs4, openpyxl, sys, re
from openpyxl.cell import *

#規定のエクセルファイルから出発駅と到着駅を取得しそれをリストに格納するファイル。
def xlsx():
    gdList = []

    ws = wb.active

    if not re.match(u'[A-Z]', from_column):
        logging.error('from_column = ({})「設定ファイルが読み込めません」'.format(from_column))
        logging.error('from_column = ({})「設定ファイルの内容が理解できません」'.format(from_column))
        sys.exit()
    elif not re.match(u'[A-Z]', to_column):
        logging.error('to_column = ({})「設定ファイルが読み込めません」'.format(to_column))
        logging.error('to_column = ({})「設定ファイルの内容が理解できません」'.format(to_column))
        sys.exit()
    elif not re.match(u'[A-Z]', price_column):
        logging.error('price_column = ({})「設定ファイルが読み込めません」'.format(price_column))
        logging.error('price_column = ({})「設定ファイルの内容が理解できません」'.format(price_column))
        sys.exit()
    elif not re.match(u'[0-9]', str(price_row)):
        logging.error('price_row = ({})「設定ファイルが読み込めません」'.format(price_row))
        logging.error('price_row = ({})「設定ファイルの内容が理解できません」'.format(price_row))
        sys.exit()
    elif not re.match(u'[0-9]', str(first_row)):
        logging.error('first_row = ({})「設定ファイルが読み込めません」'.format(first_row))
        logging.error('first_row = ({})「設定ファイルの内容が理解できません」'.format(first_row))
        sys.exit()
    elif not re.match(u'[0-9]', str(max_Row)):
        logging.error('max_Row = ({})「設定ファイルが読み込めません」'.format(max_Row))
        logging.error('max_Row = ({})「設定ファイルの内容が理解できません」'.format(max_Row))
        sys.exit()
    else:
        logging.info('「設定ファイルの読み込みが完了しました。」'.format())

        
    x = first_row
    try:        
        for i in range(1,max_Row + 1):                            #max_Row回まわす
            if sheet[from_column + str(x)].value == None:
                logging.error('sheet[from_column + str(x)].value = ({}),sheet[to_column + str(x)].value = ({})「エクセルファイルの出発駅と到着駅のどちらか、または両方が空です」,' .format(sheet[from_column + str(x)].value, sheet[to_column + str(x)].value))
                x += 1
                gdList.append('')
                gdList.append('')
                continue
            elif sheet[to_column + str(x)].value == None:
                logging.error('sheet[from_column + str(x)].value = ({}),sheet[to_column + str(x)].value = ({})「エクセルファイルの出発駅と到着駅のどちらか、または両方が空です」,' .format(sheet[from_column + str(x)].value, sheet[to_column + str(x)].value))
                x += 1
                gdList.append('')
                gdList.append('')
                continue
            elif sheet[from_column + str(x)].value == sheet[to_column + str(x)].value:
                logging.error('sheet[from_column + str(x)].value = ({}), sheet[to_column + str(x)].value = ({})「エクセルファイルの出発駅と到着駅は違うものを指定してください」'.format(sheet[from_column + str(x)].value, sheet[to_column + str(x)].value))
                x += 1
                gdList.append('')
                gdList.append('')
                continue
            elif not re.match(u'[一-龥ぁ-んァ-ン]', sheet[from_column + str(x)].value):
                logging.error('sheet[from_column + str(x)].value = ({})「出発駅と到着駅のどちらか、または両方が日本語ではありません」'.format(sheet[from_column + str(x)].value))
                x += 1
                gdList.append('')
                gdList.append('')
                continue
            elif not re.match(u'[一-龥ぁ-んァ-ン]', sheet[to_column + str(x)].value):
                logging.error('sheet[to_column + str(x)].value = ({})「出発駅と到着駅のどちらか、または両方が日本語ではありません」'.format(sheet[to_column + str(x)].value))
                x += 1
                gdList.append('')
                gdList.append('')
                continue
            else:
                gdList.append(sheet[from_column + str(x)].value)  #出発駅
                gdList.append(sheet[to_column + str(x)].value)    #到着駅
                logging.info('sheet[from_column + str(x)].value = ({}),sheet[to_column + str(x)].value = ({})「読み込んだエクセルファイルから駅名を取得しました」'.format(sheet[from_column + str(x)].value,sheet[to_column + str(x)].value))
                x += 1
                

            #一応シェルでも確認
            #print(cell_obj.coordinate, cell_obj.value)
            #print(str(getData.get(get_column_letter(from_column) + cell_obj)))
            #print('---end of row---')

        #辞書の値だけをリスト型に変換してgdListに格納
        #gdList = list(getData.values())
        #結局辞書は使わないことにした。使わないほうが設定ファイルに合わせやすいことに気付いた。
        print(gdList)
        return gdList
    except:
        logging.error('「エクセルファイルのフォーマットが異なっています。正しいフォーマットに直してください」'.format())
        sys.exit()
        




