# Student Score Counter Manual

## 파일 다운로드

1. 가상환경 생성 및 라이브러리 다운로드
    ```
    pipenv install
    ```

2. 가상환경 진입
    ```
    pipenv shell
    ```

## 프로젝트 파일 구조
프로젝트 파일 안에는 
- main.py
- 카카오톡 txt 파일(ex. kakao.txt)
- 상벌점관리 xlsx 파일(ex. 상벌점관리_01_18.xlsx)





## 프로그램 실행 방법
1. 카카오톡에서 `대화내용 내보내기` 클릭 & 다운로드
    ![Image](https://github.com/user-attachments/assets/8a4f22c9-7a59-4836-b46f-020fd25d837c)

2. 카톡 대화내용 중 필요한 부분만 추출
    ![Image](https://github.com/user-attachments/assets/3f405c81-2845-4f81-b9f8-a0eae80c41f7)
    
3. main.py의 맨 위, 파일명 및 상수를 날짜에 맞추어 수정 
    ![Image](https://github.com/user-attachments/assets/3a8b49d9-e0ae-4b29-9998-632dd4b93e18)

4. main.py 실행
 - exception이 발생한 경우, 에러가 발생한 kakao talk 줄이 출력되고, 에러 원인이 출력됨
    ![Image](https://github.com/user-attachments/assets/d65bfde5-abdc-4492-be29-682e56bbda82)
 - 에러 원인 읽어보고 kakao talk txt 파일을 명령어 규칙에 맞게 수정
 - 이후 재실행
  
5. 잘 반영되었는지 excel 파일 확인
