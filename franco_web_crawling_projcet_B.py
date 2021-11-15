
# Franko Web Crawling Project B


# Package loading


print("패키지 로딩중입니다..\n\n")

from requests import post, get
from pandas import DataFrame, read_excel
from bs4 import BeautifulSoup as bs
from numpy import array
from sys import exit


# Pre-work for Excel file output


rst_columns=[["번호","상호","영업표지","대표자","등록번호","업종","2020","","","2019","","","2018","","","2017","","","임원수","직원수","가맹사업 개시일","2020년 전체 가맹,직영점","","","2019년 전체 가맹,직영점","","","2018년 전체 가맹,직영점","","","2017년 전체 가맹,직영점","","","2020년 서울 가맹,직영점","","","2019년 서울 가맹,직영점","","","2018년 서울 가맹,직영점","","","2017년 서울 가맹,직영점","","","20년 가맹점 변동현황","","","","19년 가맹점 변동현황","","","","18년 가맹점 변동현황","","","","17년 가맹점 변동현황","","","","평균 매출액 및 면적(3.3㎡)당 매출액","","","광고, 판촉비 내역","","가맹점사업자의 부담금","","","",""],
             ["","","","","","","매출액","영업이익","당기순이익","매출액","영업이익","당기순이익","매출액","영업이익","당기순이익","매출액","영업이익","당기순이익","","","","전체","가맹점","직영점","전체","가맹점","직영점","전체","가맹점","직영점","전체","가맹점","직영점","전체","가맹점","직영점","전체","가맹점","직영점","전체","가맹점","직영점","전체","가맹점","직영점","신규개점","계약종료","계약해지","명의변경","신규개점","계약종료","계약해지","명의변경","신규개점","계약종료","계약해지","명의변경","신규개점","계약종료","계약해지","명의변경","가맹점수","평균매출액","면적당 평균매출액","광고비","판촉비","가입비","교육비","보증금","기타비용","합계"]]

df_rst = DataFrame(columns=rst_columns)
url = "https://franchise.ftc.go.kr/{}"
srch_lst_url = url.format("mnu/00013/program/userRqst/list.do")


# Loading excel files


print("'프랜차이즈 리스트' 자동 웹 크롤링 프로그램입니다.\n\n")

file_name = input("엑셀 파일의 경로를 확장자명을 포함하여 작성해 주세요.\n(예시 : C:/Users/admin/Desktop/Franko/더체크_소상공인_자료.xlsx)\n >>> ")
sheet_name = input("엑셀 시트 이름을 작성해 주세요.\n >>> ")


# If there is no data for crawling...


try:
    df = read_excel(file_name, sheet_name)
except:
    print("엑셀 파일 경로와 시트 이름을 다시 확인해 주세요.\n")
    input("엔터를 누르면 프로그램이 종료됩니다.")
    exit()
    

com_name = df["상호"].values.tolist()
com_busi_sign = df["영업표지"].values.tolist()
com_ceo = df["대표자"].values.tolist()
com_reg_num = df["등록번호"].values.tolist()
com_type = df["업종"].values.tolist()


# Data crawling


for n in range(len(com_reg_num)):
    
    print("{}번째 항목 처리 중..".format(n + 1))
    
    reg_num = com_reg_num[n]
    
    header = {
        "Referer" : srch_lst_url,
        "User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.82 Safari/537.36"
    }
    
    data = {
        "column" : "firRegNo",
        "searchKeyword" : "{}".format(reg_num),
        "selUpjong" : "",
        "selIndus" : "",
        "x" : "59",
        "y" : "16",
        "pageUnit" : "10",
    }
    
    try:
        srch_lst_res = post(srch_lst_url, headers = header, data = data).text
    except:
        print("페이지가 일시적으로 다운되었거나 인터넷 연결이 끊어졌습니다. 잠시 후 다시 실행해 주세요.\n")
        input("엔터를 누르면 프로그램이 종료됩니다.")
        exit()
        
    srch_lst_soup = bs(srch_lst_res, "html.parser")


    try:
        srch_com = srch_lst_soup.select("#frm > table > tbody > tr > td:nth-child(2) > a")[0]["onclick"][10:].replace("');" , "")
    except IndexError:
        print("'{}' 해당 상호는 가맹사업정보제공시스템에 존재하지 않습니다. 다음 항목으로 넘어갑니다.".format(com_name[n]))
        continue
    
    srch_com_url = url.format(srch_com)
    srch_com_res = get(srch_com_url).text
    srch_com_soup = bs(srch_com_res, "html.parser")
    
    is_there_2020 = srch_com_soup.select("#frm > div:nth-child(13) > div > table:nth-child(4) > thead > tr:nth-child(1) > th:nth-child(2)")[0].text.split()[0]
    
    if is_there_2020 == "2020년":
        is_there_2020 = True
    else:
        is_there_2020 = False
    
    head_fin = []
    fran_emplo = []
    fran_open_date = ""
    store_num_all = []
    store_num_seoul = []
    fran_changed = []
    area_sales = []
    ad_cost = []
    fran_allo = []
    
    if is_there_2020 == False:
        for i in range(3):
            head_fin.append("")
            store_num_all.append("")
            store_num_seoul.append("")
        for i in range(4):
            fran_changed.append("")
    
    for i in range(2,5):
        for j in range(5,8):
            try:
                com_data = srch_com_soup.select("#frm > div:nth-child(12) > div > table:nth-child(5) > tbody > tr:nth-child({0}) > td:nth-child({1})".format(i , j))[0].text
                if com_data == "":
                    raise IndexError
                head_fin.append(com_data)
            except IndexError:
                print("'{}' 해당 상호의 재무상황에 누락된 데이터가 있습니다.".format(com_name[n]))
                head_fin.append("")
                continue

    for i in range(2,4):
        try:
            com_data = srch_com_soup.select("#frm > div:nth-child(12) > div > table:nth-child(7) > tr > td:nth-child({})".format(i))[0].text
            if com_data == "":
                raise IndexError
            fran_emplo.append(com_data)
        except IndexError:
            print("'{}' 해당 상호의 임직원수에 누락된 데이터가 있습니다.".format(com_name[n]))
            fran_emplo.append("")
            continue
    
    try:
        fran_open_date = srch_com_soup.select("#frm > div:nth-child(13) > div > table:nth-child(2) > tbody > tr > td")[0].text.split()
        if fran_open_date == "":
            raise IndexError
        try:
            fran_open_date = int(fran_open_date[0])
        except ValueError:
            fran_open_date = " ".join(fran_open_date)
            pass
    except IndexError:
        print("'{}' 해당 상호의 개시일에 누락된 데이터가 있습니다.".format(com_name[n]))
        fran_open_date = ""
        continue
    
    for i in range(2, 11):
        try:
            com_data = srch_com_soup.select("#frm > div:nth-child(13) > div > table:nth-child(4) > tbody > tr:nth-child(1) > td:nth-child({0})".format(i))[0].text.split()[0]
            if com_data == "":
                raise IndexError            
            store_num_all.append(com_data)
        except IndexError:
            print("'{}' 해당 상호의 가맹점 및 직영점 현황(전체)에 누락된 데이터가 있습니다.".format(com_name[n]))
            store_num_all.append("")
            continue

    for i in range(2, 11):
        try:
            com_data = srch_com_soup.select("#frm > div:nth-child(13) > div > table:nth-child(4) > tbody > tr:nth-child(2) > td:nth-child({0})".format(i))[0].text.split()[0]
            if com_data == "":
                raise IndexError            
            store_num_seoul.append(com_data)
        except IndexError:
            print("'{}' 해당 상호의 가맹점 및 직영점 현황(서울)에 누락된 데이터가 있습니다.".format(com_name[n]))
            store_num_seoul.append("")
            continue
    
    for i in range(0, 3):
        for j in range(2, 6):
            try:
                com_data = srch_com_soup.select("#frm > div:nth-child(13) > div > table:nth-child(6) > tr > td:nth-child({0})".format(j))[i].text.split()[0]
                if com_data == "":
                    raise IndexError
                fran_changed.append(com_data)
            except IndexError:
                print("'{}' 해당 상호의 가맹점 변동 현황에 누락된 데이터가 있습니다.".format(com_name[n]))
                fran_changed.append("")
                continue
        
    for i in range(2, 5):
        try:
            com_data = srch_com_soup.select("#frm > div:nth-child(13) > div > table:nth-child(8) > tbody > tr:nth-child(1) > td:nth-child({0})".format(i))[0].text.split()[0]
            if com_data == "":
                raise IndexError
            area_sales.append(com_data)
        except IndexError:
            print("'{}' 해당 상호의 평균 매출액 및 면적 당 평균매출액에 누락된 데이터가 있습니다.".format(com_name[n]))
            area_sales.append("")
            continue
    
    for i in range(2, 4):
        try:
            com_data = srch_com_soup.select("#frm > div:nth-child(13) > div > table:nth-child(12) > tbody > tr > td:nth-child({})".format(i))[0].text.split()[0]
            if com_data == "":
                raise IndexError
            ad_cost.append(com_data)
        except IndexError:
            print("'{}' 해당 상호의 광고 및 판촉비 내역에 누락된 데이터가 있습니다.".format(com_name[n]))
            ad_cost.append("")
            continue
        
    for i in range(1, 6):
        try:
            com_data = srch_com_soup.select("#frm > div:nth-child(15) > div > table:nth-child(2) > tbody > tr > td:nth-child({})".format(i))[0].text.split()[0]
            if com_data == "":
                raise IndexError
            fran_allo.append(com_data)
        except IndexError:
            print("'{}' 해당 상호의 가맹점사업자의 부담금에 누락된 데이터가 있습니다.".format(com_name[n]))
            fran_allo.append("")
            continue
    
    if is_there_2020 == True:
        for i in range(3):
            head_fin.append("")
            store_num_all.append("")
            store_num_seoul.append("")
        for i in range(4):
            fran_changed.append("")

            
# Storing data to Excel file

    rst_lst = []
    
    rst_lst.append(n + 1)
    rst_lst.append(com_name[n])
    rst_lst.append(com_busi_sign[n])
    rst_lst.append(com_ceo[n])
    rst_lst.append(com_reg_num[n])
    rst_lst.append(com_type[n])
    rst_lst = rst_lst + head_fin + fran_emplo
    rst_lst.append(fran_open_date)
    rst_lst = rst_lst + store_num_all + store_num_seoul + fran_changed + area_sales + ad_cost + fran_allo
    
    
    df_item_rst = DataFrame(data=array([rst_lst]),columns=rst_columns)
    df_rst = df_rst.append(df_item_rst)

rst_file = sheet_name + "_" + 'result.xlsx'

print("\n결과물이 해당 프로그램이 있는 디렉토리에 '{}'로 저장되었습니다.".format(rst_file))
df_rst.to_excel(rst_file)

input("\n엔터를 누르면 프로그램이 종료됩니다.")
exit()

# Created by jihyun jung