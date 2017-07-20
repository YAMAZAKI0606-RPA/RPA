from op import *
from selenium import webdriver
from xlsx import *
import time, bs4, openpyxl, sys, requests, lxml, urllib.parse



#gdListに格納された値をそれぞれの入力フォームに入れる
def web():
    gdList = xlsx()
    print(str(gdList) + 'web.pyのgdListです')
    stup = []
    print(str(len(gdList)) + 'wbのgdListの長さです')


    try:   
        for i in range(0,len(gdList),2):
            try:
                #chromeでyahoo乗り換えページを表示
                browser = webdriver.Chrome()
                browser.get('https://transit.yahoo.co.jp/')
                logging.info('「ブラウザの表示に成功しました」')
            except:
                logging.error('「インターネットに接続できません」'.format())
                sys.exit()

            
            #出発駅を取得し入力
            if gdList[i] == '':
                logging.error('gdList[i] = ({}),gdList[i + 1] = ({})「出発駅の入力フォームと到着駅の入力フォームのどちらか、または両方が入力出来ません」'.format(gdList[i],gdList[i + 1]))
                logging.error('gdList[i] = ({}),gdList[i + 1] = ({})「ブラウザを閉じ、次の行に飛びます」'.format(gdList[i],gdList[i + 1]))
                browser.close()
                stup.append('')
                continue
            else:
                from_t = browser.find_element_by_id('sfrom')
                from_t.send_keys(gdList[i])
                print(gdList[i])
                logging.info('gdList[i] = ({})「出発駅の入力に成功しました」'.format(gdList[i]))
            

            #到着駅を取得し入力
            if gdList[i] == '':
                logging.error('gdList[i] = ({}),gdList[i + 1] = ({})「出発駅の入力フォームと到着駅の入力フォームのどちらか、または両方が入力出来ません」'.format(gdList[i],gdList[i + 1]))
                logging.error('gdList[i] = ({}),gdList[i + 1] = ({})「ブラウザを閉じ、次の行に飛びます」'.format(gdList[i],gdList[i + 1]))
                browser.close()
                stup.append('')
                continue
            else:
                to_t = browser.find_element_by_id('sto')
                i += 1
                to_t.send_keys(gdList[i])
                print(gdList[i])
                logging.info('gdList[i] = ({})「到着駅の入力に成功しました」'.format(gdList[i]))
            
            #送信(submit)
            from_t.submit()

            #念のため待機時間をセット
            time.sleep(1)

            try:
                #最安値をクリック
                link_elem = browser.find_element_by_link_text('[安]料金の安い順')
                type(link_elem)
                link_elem.click()
                logging.info('「「料金の安い順」へのクリックが成功しました」')
            except:
                logging.error('gdList[i] = ({}),gdList[i + 1] = ({})「フォームに入力された駅名から最安値を検索出来ませんでした」'.format(gdList[i],gdList[i + 1]))
                logging.error('gdList[i] = ({}),gdList[i + 1] = ({})「料金の安い順をクリック出来ませんでした」'.format(gdList[i],gdList[i + 1]))
                stup.append('')
                browser.close()
                continue




            #念のため待機時間をセット
            time.sleep(1)

            #webページを閉じる
            #browser.quit()

            #ブラウザを閉じる(この時エラーメッセージが出るがどう消すかは未知。消さないとどんどん溜まっていく。エラーを消す方法が見つからないのでブラウザを閉じずに
            #参考ログとして残すことにする。)!!!!!!!!クロームだとプログラム上からブラウザを閉じてもエラーが出ない！！びっくり！！
            #browser.close()

            #最安値のページを開き、最安値の情報を取得する。
            try:
                #raise Exception('最安値取得失敗')
                print(browser.current_url)
                res = requests.get(browser.current_url)
                res.raise_for_status()
                lowcost = bs4.BeautifulSoup(res.text)
                type(lowcost)
                #指定の仕方が甘く二回目以降は到着時間が表示されてしまう。まだ未修整。!!!!複数回表示されるspanとはだめでHTMLの固有のタグで囲まれたものならいけることに気づいた。!!!!
                elems = lowcost.select('li.fare')
                stup.append(elems[0].getText())
                type(elems)
                print(elems[0].getText())
                print(str(stup) + 'web.pyのstupです')
                time.sleep(1)
                browser.close()
                logging.info('elems[0].getText() = ({})「最安値の取得に成功しました」'.format(elems[0].getText()))
            except:
                logging.error('gdList[i] = ({}),gdList[i + 1] = ({})「最安値を取得出来ませんでした」'.format(gdList[i],gdList[i + 1]))
                stup.append('')
                browser.close()
                continue
        return stup
    except:
        logging.error('「インターネットの接続が途切れました。再度接続しなおしてください」'.format())
        sys.exit()
        
          




        #webページを閉じる
        #browser.quit()

        #ブラウザを閉じる(この時エラーメッセージが出るがどう消すかは未知。消さないとどんどん溜まっていく)
        #browser.close()
