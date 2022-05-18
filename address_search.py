import urllib.request
import json
import requests
import pandas as pd
import numpy as np
import openpyxl
import re

def openxl(file_name,name_colum, address_colum):
    wb = openpyxl.load_workbook('C:\\search\\'+file_name+'.xlsx',data_only=True)
    ws = wb.active
    exel_df = pd.DataFrame(ws.values)

    exel_df = exel_df.loc[:,[name_colum-1,address_colum-1]]
    exel_df = exel_df.drop([0])
    exel_df.columns = ['Name','Address']
    print(exel_df)
    return exel_df

def name_preproccesing(name):
    change = re.sub('\([^\)]*\)|[^\w\s]',"",name)

    return change


# 카카오맵 상호명 검색
def name_search(name):
    url = 'https://dapi.kakao.com/v2/local/search/keyword.json'
    params = {'query': '인천 서구 '+name}
    headers = {
        "Authorization": "KakaoAK 869d116685514b64b61f5e99ad0a5059"
    }
    places = requests.get(url, params=params, headers = headers).json()['documents']
    return places
'''
# 카카오맵 주소 검색
def address_search(address):
    url = 'https://dapi.kakao.com/v2/local/search/address.json'
    params = {'query': address}
    headers = {
        "Authorization": "KakaoAK 869d116685514b64b61f5e99ad0a5059"
    }
    places = requests.get(url, params=params, headers = headers).json()['documents']
    print(places)
    return places

address_search('인천 서구 중봉대로 610')
'''


# 검색된 정보 Df저장
def info(places):
    stores = []
    road_address = []

    for place in places:
        stores.append(place['place_name'])
        road_address.append(place['road_address_name'])

    arr = np.array([stores,road_address]).T
    df = pd.DataFrame(arr, columns=['Name','address'])
    print(df)
    return df

def naver_search(name,address):
    client_id = "jbSZRInk6ZW_BHIo9qRT"
    client_secret = "8z_JIqwEaT"
    encText = urllib.parse.quote("인천광역시 서구 "+name)
    display = "&display=5"
    url = "https://openapi.naver.com/v1/search/local.json?query=" + encText +display # json 결과
    # url = "https://openapi.naver.com/v1/search/blog.xml?query=" + encText # xml 결과
    request = urllib.request.Request(url)
    request.add_header("X-Naver-Client-Id",client_id)
    request.add_header("X-Naver-Client-Secret",client_secret)
    response = urllib.request.urlopen(request)
    #print(response)

    address_list =[]
    rescode = response.getcode()
    if(rescode==200):
        response_body = response.read()

        df = response_body.decode('utf-8')

        data= json.loads(df)
    else:
        print("Error Code:" + rescode)

    dd =data['items']
    result = False
    for i in dd:
        addr = i['roadAddress']
        #print ('네이버 주소검색: '+i['roadAddress'])
        address_list.append(addr)

    for i in address_list:

        if address in i:
            print("naver 검색: "+i)
            result = True
            break

        else: result = False

    return result



#기존 정보와 검색된 정보 비교
def comparison(places):
    comp = []
    comp_value = False
    road =" "
    for name,address in zip(places['Name'],places['Address']):
        try:
            pre_name = name_preproccesing(name)
            search_loc = name_search(pre_name)

            info_place_df = info(search_loc)

            de = re.sub('\,.*|\(.*|\..*|[A-Z].*',"",address)
            p = re.compile('\S+로.*\d')


            road_address =(p.findall(de))
            print(pre_name)
            print(road_address)

        except TypeError:
            road_address =["none"]
        #    print(road_address)
        if (not info_place_df.empty):
            for add in info_place_df['address']:
                road = add
                try:
                    if (road_address[0] in add):

                        comp_value = True
                        break
                    elif naver_search(pre_name,road_address[0]):
                        comp_value = True
                        break
                    else: comp_value = False

                except IndexError:
                    comp_value = False
                except TypeError:
                    comp_value = False
                except Exception as ex:
                    print(ex+' 에러가 발생 했습니다.')
                    comp_value = False

        else: comp_value = False

        print("검색된 도로명 주소 : "+road)
        if comp_value:
            comp.append('O')
            print('검색결과 : O')
        else:
            comp.append('X')
            print('검색결과 : X')


    print(comp)

    return comp

# 파일 불러오기 (파일명, 상호명 열위치, 도로주소 열위치)
file_name = '134.인천광역시_서구_건강기능식품판매업소_20210702'
name_colums = 3
address_colums = 4

place =pd.DataFrame()
place = openxl(file_name,name_colums,address_colums)

'''
writer = pd.ExcelWriter(file_name+'_주소확인.xlsx')
place.to_exel(file_name+'_주소확인.xlsx')
'''

# 검색 후 주소 비교
comparison_result = comparison(place)

# 기존 테이블에 검색결과 추가
place['Result'] = comparison_result


print(place)

# 검색실패한 주소
place_none= place.drop(['Address'],axis=1)
x_value= place['Result']
place_none = place_none[place_none['Result']=='X']


print(place_none)
print(x_value)

place.to_csv(file_name+'_(ox)검색완료.csv',sep="\t",encoding='cp949')

place_none.to_csv(file_name+'_검색실패.csv',sep="\t",encoding='cp949')
#x_value.to_csv(file_name+'_x.csv',sep="\t",encoding='euc_kr',index=False)
#writer = pd.ExcelWriter(file_name+'_주소확인.xlsx')


