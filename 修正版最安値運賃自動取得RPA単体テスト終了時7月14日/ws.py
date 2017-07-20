# -*- coding: utf-8 -*-
import wx, openpyxl, types, sys
from op import *

#実行の際エクセルファイルの選択ダイアログを表示させるファイル。変数受け渡しが上手くいかなかったせいで別ファイルにせざるを得なかった。他にいい方法があれば知りたい。
app = wx.App()

# ファイル選択ダイアログを作成。ファイル選択の際拡張子で判断したしたいが、なぜかALLになってしまう。
#filter = 'Excel ファイル(*.xlsx) | *.xlsx | *.*'
#filter = 'Python files (*.py;*.pyw) | *.py;*.pyw'
try:
    dialog = wx.FileDialog(None, u'エクセルファイルを選択してください')

# ファイル選択ダイアログを表示
    dialog.ShowModal()
    logging.info('「エクセルファイル選択ログの表示に成功しました」')
except:
    logging.error('「ファイル選択ダイアログを開けませんでした」'.format())
    logging.error('「エクセルファイルが読み込めません」'.format())
    sys.exit()


#dlg = wx.FileDialog(self, "ファイルを選択してください", self.dirname, "", "*.*", wx.OPEN)

print(dialog.GetFilename())
AAA = dialog.GetFilename()
if 'xlsx' not in AAA:
    logging.error('AAA({})「エクセルファイルが読み込めません。拡張子を.xlsxに合わせてください」'.format(AAA))
    sys.exit()
else:
    logging.info('AAA({})「選択したファイルの拡張子は.xlsxです」'.format(AAA))

name = dialog.GetFilename()
logging.info('name = {}' .format(name))

#エクセルシートオブジェクト取得
try:
    wb = openpyxl.load_workbook(name)
    logging.info('「エクセルファイルオブジェクトの取得が完了しました」'.format())
    #logging.info('name = ({})エクセルファイルが指定されました,' .format(name))

#if type(wb) is None:
except:
    logging.error('「エクセルファイルが読み込めません。指定したファイルの内容を確認してください」'.format())
    logging.error('「ファイル選択ダイアログを開けませんでした」'.format())
    sys.exit()
    
#else:
    #logging.info('wb = ({})エクセルファイルが指定されました。'.format(wb))

sheet = wb.get_sheet_by_name('Sheet1')
tuple(sheet['A3':'C3'])

    #一応確認
print('横の行は：'+str(sheet.max_row)+'   '+'縦の列は：'+ str(sheet.max_column))



