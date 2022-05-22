# Data_Internship_OfficeAutomation
공공데이터 인턴 Naver, Kakao Map을 이용한 주소비교 업무자동화


## ⚙️process
1.  Excel or CSV 파일명과 찾을 **상호 or 회사명, 도로명 주소의 열 번호를 입력**
![image](https://user-images.githubusercontent.com/57780594/169347209-a8598eeb-a29e-4484-97de-2d436c4f483e.png)
<br>

2.  해당 파일의 상호 or 회사명과 도로명 주소만 불러와 **데이터프레임에 저장**
<br>  
  
3.  **Kakao Map Api**를 이용해 상호 or 회사명으로 검색 후 주소 저장
<br>
 
4.  검색 실패 시 **Naver Search Api**를 이용해 상호 or 회사명으로 검색 후 주소 저장
<br>
  
5.  검색된 주소와 기존 공공데이터 파일의 주소를 **비교 전 전처리 진행**  
![image](https://user-images.githubusercontent.com/57780594/169348652-a33bb144-d704-439d-9e6d-a3d80a8e2cc0.png)
<br>
  
6. **비교 시 같다면 O, 틀리면 X** 값이 저장된 열을 데이터프레임에 추가
![image](https://user-images.githubusercontent.com/57780594/169351155-199c9bf8-51e9-4d6b-a946-15903f22c243.png)
<br>

7. 해당 데이터프레임을 다시 Excel or CSV로 저장
