import csv
import getpass
import openpyxl
import os
import re
import sys
from selenium.webdriver import Chrome, ChromeOptions

# UNIPAのログイン情報を取得
print('UNIPAにログインします。ユーザー情報を入力して下さい。')
user = input('ユーザー名: ')
if sys.stdin.isatty():
    password = getpass.getpass('共通パスワード: ')
else:
    print('Using readline')
    password = sys.stdin.readline().rstrip()

options = ChromeOptions()
# ヘッドレスモードを有効にする(次の行をコメントアウトすると画面が表示される)
options.add_argument('--headless')
# ChromeのWebDriverオブジェクトを作成する
driver = Chrome(executable_path='./drivers/chromedriver', options=options)

# スクレイピング開始
print('>>> Chromeヘッドレスブラウザ 起動')
driver.get('https://portal.sa.dendai.ac.jp/up/faces/login/Com00505A.jsp')  # UNIPAのログインページURLへ移動

print('>>> ログインページ画面 操作中')
input_username = driver.find_element_by_id('form1:htmlUserId')  # 「User ID」のinput
input_password = driver.find_element_by_id('form1:htmlPassword')  # 「PassWord」のinput
input_login = driver.find_element_by_id('form1:login')  # 「ログイン」ボタンのinput

input_username.send_keys(user)  # 入力したユーザー名をinputに記述
input_password.send_keys(password)  # 入力したパスワードをinputに記述
input_login.click()  # ログインボタンをクリック

print('>>> ホーム画面 操作中')
# 成績照会のonclickイベントが上手く動作しないので、onclickイベントの内容のjavascriptを直接実行
driver.execute_script('javascript:clickMenuItem(602,0)')

print('>>> 成績照会画面 操作中')
credit_table = driver.find_element_by_class_name('singleTableLine')  # 成績照会の成績全体のテーブル

subjects = credit_table.find_elements_by_class_name('tdKamokuList')  # 科目名
years = credit_table.find_elements_by_class_name('tdNendoList')  # 履修年度
credits = credit_table.find_elements_by_class_name('tdTaniList')  # 単位数
assessments = credit_table.find_elements_by_class_name('tdHyokaList')  # 評価

# 結果を格納する変数を用意
group = ''  # 科目分類を格納する変数
got_credits = []  # 履修済み単位を格納する変数
taking_courses = []  # 履修中単位を格納する変数

# スクレイピングの結果の値を変数に格納
for num in range(len(subjects)):
    # 科目名が科目分類になってた時
    if re.match(r'＜(他部科)?(.)*(必修)?(選択)?科目＞', subjects[num].text):
        # 大なり小なりとかの文字が邪魔なので切り取る
        group = re.sub('＜|他部科|必修|選択|科目|＞', '', subjects[num].text)
    else:
        # 履修中の科目(評価の欄が空欄)の時
        if assessments[num].text == '':
            # 評価以外を変数に格納
            taking_courses.append([subjects[num].text, years[num].text, credits[num].text, group])
        # 履修済み科目の時
        else:
            # すべてを変数に格納
            got_credits.append([subjects[num].text, years[num].text, credits[num].text, assessments[num].text, group])

print('>>> Chromeヘッドレスブラウザ 終了')
# ChromeのWebDriverオブジェクトを正常に終了する
driver.quit()

# csvを書き出す場所を指定
results_dir = './csv_results/'
# 書き出す場所が無い時
if not os.path.exists(results_dir):
    # ディレクトリを作る
    os.mkdir(results_dir)

# 履修済み科目のリストをcsvに書き出す
print('>>> CSVファイル書き出し中')
with open(results_dir + 'got_credits.csv', 'w') as got_credits_csv:  # デフォルトのエンコードはUTF-8
    writer = csv.writer(got_credits_csv, lineterminator='\r\n')
    for num in range(len(got_credits)):
        writer.writerow(got_credits[num])  # データをcsvに書き込み

# 履修中科目のリストをcsvに書き出す
with open(results_dir + 'taking_courses.csv', 'w') as taking_courses_csv:  # デフォルトのエンコードはUTF-8
    writer = csv.writer(taking_courses_csv, lineterminator='\r\n')
    for num in range(len(taking_courses)):
        writer.writerow(taking_courses[num])  # データをcsvに書き込み

# Excelファイルの2つのシートに追記
print('>>> Excelファイル更新中')
wb = openpyxl.load_workbook('recode.xlsx')
got_credits_sheet = wb['履修済み一覧']
taking_courses_sheet = wb['履修予定一覧']

# 履修済み科目リストを '履修済み一覧' シートに書き込み
for num in range(len(got_credits)):
    got_credits_sheet['A' + str(num + 2)].value = got_credits[num][0]
    got_credits_sheet['B' + str(num + 2)].value = got_credits[num][1]
    if got_credits[num][2] == '':
        got_credits_sheet['C' + str(num + 2)].value = got_credits[num][2]
    else:
        got_credits_sheet['C' + str(num + 2)].value = float(got_credits[num][2])
    got_credits_sheet['D' + str(num + 2)].value = got_credits[num][3]
    got_credits_sheet['E' + str(num + 2)].value = got_credits[num][4]

# 履修中科目リストを '履修予定一覧' シートに書き込み
for num in range(len(taking_courses)):
    taking_courses_sheet['A' + str(num + 2)].value = taking_courses[num][0]
    taking_courses_sheet['B' + str(num + 2)].value = taking_courses[num][1]
    taking_courses_sheet['C' + str(num + 2)].value = float(taking_courses[num][2])
    taking_courses_sheet['D' + str(num + 2)].value = taking_courses[num][3]

# Excelファイルを上書き保存
wb.save('recode.xlsx')
