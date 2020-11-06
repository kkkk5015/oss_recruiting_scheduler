import webbrowser
import openpyxl as xl
import datetime as d
import requests
from bs4 import BeautifulSoup
import urllib
from openpyxl.workbook import Workbook
from openpyxl.styles import Font

'''
    기능 저장, 분류, ...

    첫화면 기능 선택
        저장 시 필요한 요소
            회사명, 마감일, 코테 유무(날짜), 회사지원url, 코테 지원 언어,...

        분류
            지원 날짜별,
'''

def create(): #엑셀파일 만들어서 구분하기
    wb = Workbook()
    ws = wb.active
    ws['A1'] = "회사명"
    # 폰트 스타일링
    ws['A1'].font = Font(bold=True,color='000000FF',size=15)
    ws['B1'] = "마감일"
    ws['B1'].font = Font(bold=True,color='000000FF',size=15)
    ws['C1'] = "코테 유무(날짜)"
    ws['C1'].font = Font(bold=True,color='000000FF',size=15)
    ws['D1'] = '회사지원url'
    ws['D1'].font = Font(bold=True,color='000000FF',size=15)
    ws['E1'] = '코테 지원 언어'
    ws['E1'].font = Font(bold=True,color='000000FF',size=15)
    # ht = input("엑셀파일 만들기 위해 저장할 위치 입력해주세요 : ")
    wb.save('Company Schedules.xlsx')
    print("현재 파일 위치에 'Company Schedules'라는 엑셀 파일을 만들었습니다")

def xlsave():  # 액셀 저장 파트
    create()
    while True:
        print()
        ent = []
        company_name = str(input("지원 하실 회사의 이름을 적어주세요(첫 화면으로 돌아가려면 *을 적어주제요) : "))
        if company_name == '*':
            break
        d_day = str(input("지원 마감일을 적어주세요(예시 : 09/24) : "))
        code_day = str(input("코딩테스트 날짜를 적어주세요(예시: 09/27, 코딩 테스트가 없다면 x입력) : "))
        com_url = str(input("회사 지원 홈페이지 주소를 적어주세요 : "))
        code_lan = str(input("코딩테스트 언어를 모두 적어 주세요, 테스트가 없다면 x입력 (예시 : java, c++, python) : "))
        ent = [company_name, d_day, code_day, com_url, code_lan]
        wb = xl.load_workbook('Company Schedules.xlsx')
        ws = wb['Sheet']
        # 혹시 필요하다면 sheet 값도 받아야함
        ws.append(ent)
        wb.save('Company Schedules.xlsx')
        print()
        print("지원 회사가 저장되었습니다. ")
        print()


def xlout():  # 조회 기능 파트
    # 기능 지원 회사 나열, ㅇ
    print("1. ")
    print("2. 회사 정보 조회하기")
    print("3. 회사 지원홈페이지 가기")
    print("4. 지원 마감된 제외하고 회사 나열하기")
    print("5. 코딩테스트가 있는 지원회사 나열하기")
    print("")
    a = input("원하시는 업무 번호를 적어주세요(나가기는 * 입력) : ")


def xlrecurl() :
    print()
    # ht = input("저장된 엑셀파일의 주소를 적어주세요 : ")
    wb = xl.load_workbook('Company Schedules.xlsx')
    # 필요하다면 sheet 값도 받아야한다.
    ws = wb['Sheet']
    col_value = []
    for col in ws.columns:
        col_l = []
        for cell in col:
            col_l.append(cell.value)
        col_value.append(col_l)
    print()
    for i, com in enumerate(col_value[0][1:]):
        print('{}. {}'.format(i, com))
    print()
    comnum = int(input("원하는 회사의 번호를 입력해주세요 : "))
    print()
    comurl = col_value[3][comnum]
    webbrowser.open(comurl)


def xlcominfo():  # 회사 정보 가져오기
    print()
    comname = input("조회할 기업명을 입력해주세요 : ")
    cominfo = []
    temp = []
    # comname = '파수닷컴'
    cominfourl = 'http://www.jobkorea.co.kr/Salary/Index?coKeyword=' + comname + '&tabindex=0&indsCtgrCode=&indsCode=&jobTypeCode=&haveAGI=0&orderCode=2&coPage=1#salarySearchCompany'
    # mykey = {'keyword' : comname }
    r = requests.get(cominfourl)
    soup = BeautifulSoup(r.text, 'html.parser')
    keyword = soup.select('.salaryCompanyList .container .list li a')
    # print(keyword)
    for word in keyword:
        w = list(map(str, word.get_text().split('\n')))
        temp.append(w)

        # temp = list(map(str, word.get_text().split('\n')))
    # print(temp)
    for c in temp:
        t = []
        for k in c:
            if k in t or k == '':
                continue
            elif k == '좋아요' or k == '채용중':
                continue
            t.append(k)

        cominfo.append(t)

    for com in cominfo:
        for p, line in enumerate(com):
            print(line)
            if p == 0:
                print()
        print()
        print('-------------------------------')
        print()

    '''
    print("기업명 : "+cominfo[0])
    print("회사직무 : "+cominfo[3])
    print(cominfo[4])
    print(cominfo[5])
    print("평균연봉 : "+cominfo[6])
    '''


while True:
    print()
    print("1. 지원 일정 저장하기")
    print("2. 지원 일정 조회하기")
    print("원하시는 업무 번호를 적어주세요(나가기는 * 입력) : ", sep="", end="")
    st = input()
    if st == '*':
        break
    if st == '1':
        xlsave()
    if st == '2':
        xlout()
