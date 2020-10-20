import openpyxl as xl
import datetime as d
'''
    기능 저장, 분류, ...
    
    첫화면 기능 선택 
        저장 시 필요한 요소 
            회사명, 마감일, 코테 유무(날짜), 회사지원url, 코테 지원 언어,...
            
        분류 
            지원 날짜별,  
'''

def xlsave() :
    #print("일정을 저장할 엑셀파일의 주소를 적어주세요 : ",sep='',end='')
    ht = input("일정을 저장할 엑셀파일의 주소를 적어주세요 : ")
    while True :
        ent = []
        company_name = str(input("지원 하실 회사의 이름을 적어주세요(첫 화면으로 돌아가려면 *을 적어주제요) : "))
        if company_name == '*' :
            break
        d_day = str(input("지원 마감일을 적어주세요(예시 : 09/24) : "))
        code_day = str(input("코딩테스트 날짜를 적어주세요(예시: 09/27, 코딩 테스트가 없다면 x입력) : "))
        com_url = str(input("회사 지원 홈페이지 주소를 적어주세요 : "))
        code_lan = str(input("코딩테스트 언어를 모두 적어 주세요, 테스트가 없다면 x입력 (예시 : java, c++, python) : "))
        ent = [company_name,d_day,code_day,com_url,code_lan]
        wb = xl.load_workbook(ht)
        ws = wb['Sheet']
        ws.append(ent)
        wb.save(ht)
        print()
        print("지원 회사가 저장되었습니다. ")
        print()






while True :
    print("1. 지원 일정 저장하기")
    print("2. 지원 일정 조회하기")
    print("원하시는 업무 번호를 적어주세요(나가기는 * 입력) : ", sep="",end="")
    st = input()
    if st == '*' :
        break
    if st == '1' :
        xlsave()







