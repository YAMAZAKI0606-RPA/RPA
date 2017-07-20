import logging , datetime

#now = timedatetime().now
d = datetime.datetime.today()
logging.basicConfig(filename = 'log_RPA' + "{0:%Y%m%d_%H%M%S}.txt".format(d), level = logging.DEBUG,
format = ' %(asctime)s - %(levelname)s - %(message)s')



#ログレベルはここで調節できる。コメントアウトすればDEBUGレベルから表示される。
logging.disable(logging.DEBUG)




#この列の変数を変えることで、それぞれの列を横に移動しても対応できる。他の動かし方は未検証。


#出発駅の列です。イコール(=)の後に半角スペースを空けて、半角大文字アルファベットで入力し'で囲むのを忘れないでください。
from_column = 'A'




#到着駅の列です。イコール(=)の後に半角スペースを空けて、半角大文字アルファベットで入力し'で囲むのを忘れないでください。
to_column = 'C'




#最安値が入力される列です。イコール(=)の後に半角スペースを空けて、半角大文字アルファベットで入力し'で囲むのを忘れないでください。
price_column = 'E'




#最安値が入力される最初の行です。イコール(=)の後に半角スペースを空けて、半角数字で入力してください。
price_row = 3




#読み込まれる最初の行です。イコール(=)の後に半角スペースを空けて、半角数字で入力してください。
first_row = 3




#読み込む限界の行です。イコール(=)の後に半角スペースを空けて、半角数字で入力し、表題や空欄は含まずに駅名が実際に入っているセルの行数を入力してください。
max_Row = 8


