from op import *
from ws import *
from web import *
from xlsx import *
import openpyxl, sys


#取得した値(最安値)をエクセルに入力するメソッド。最後に元のエクセルファイルに上書きするか、新しいファイルを作成し保存するかは未定(現時点では上書き)。


def xlsx_In():
    
    #stupに最安値を格納
    stup = web()
    p = price_row
    e = 0
    print(str(stup) + 'stupの中身です')
    for i in range(max_Row):
        print(i)
        print(str(stup[i]))
        st = sheet[price_column + str(p)].value

        #セルに価がなく、最安値がある
        if st == None and stup[i] != '':
            sheet[price_column + str(p)].value = stup[i]
            logging.info('sheet[price_column + str(p)].value = ({})「エクセルファイルに取得した最安値の入力に成功しました」'.format(sheet[price_column + str(p)].value))
            p += 1
            i += 1
        #セルに価がなく、最安値がない
        elif st == None and stup[i] == '':
            logging.error('stup[i] = ({})「最安値を取得していないのでエクセルファイルに入力出来ません」'.format(stup[i]))
            e += 1
            p += 1
            i += 1
        #セルに価があり、最安値がある
        elif st != None and stup[i] != '':
            logging.error('p = ({}),i = ({})「最安値は取得できましたが、エクセルファイルの最安値の列に既に値が入っているので上書きできません」'.format(p,i))
            e += 1
            p += 1
            i += 1
        #セルに価があり、最安値がない
        else:
            logging.error('p = ({}),i = ({})「最安値は取得できず、エクセルファイルの最安値の列に既に値が入っているので上書きできません」'.format(p,i))
            e += 1
            p += 1
            i += 1           
    
    if e == 0:
        wb.save(name)
        print('エクセルに値を登録しました')
        logging.info('「最安値の入力、保存に成功しました」')
        sys.exit()
    elif e > 0:
        wb.save(name)
        print(str(i) + '件中' + str(e) + '件がエクセルファイルへの入力に失敗しました')
        logging.error('i = ({})件中e = ({})件がエクセルファイルへの入力に失敗しました」'.format(i,e))
        sys.exit()
    else:
        logging.error('「入力予定のエクセルファイルが開かれています。閉じてから再度実行してください」'.format())
        sys.exit()



