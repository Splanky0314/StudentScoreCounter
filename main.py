from openpyxl import load_workbook # excel lib
from openpyxl.comments import Comment
from functions import *
from student_data import *


# 수정 필요 변수
TARGET_COL = 21 # 날짜 기반 
search_row_range = 166
AUTHOR = "김다은"
# EXCEL_FILE_NAME = "상벌점관리_01_14.xlsx"
EXCEL_FILE_NAME = "test.xlsx"
KAKAO_TXT_NAME = "kakao2.txt"

# excel 
wb = load_workbook(EXCEL_FILE_NAME)
ws = wb.active

# kakao txt 읽어오기
f = open(KAKAO_TXT_NAME, 'rt', encoding='UTF8')

lines = f.readlines()
for line in lines:
    if "ㅇㅇ" in line:
        """
        << 한명 상벌점(ㅇㅇ) 처리 >>
        ㅇㅇ 홍길동 +3 수학태도우수
        ㅇㅇ 홍길동 -3 수학태도불량
        """
        response = find_command_and_return_index(line, "ㅇㅇ")
        stdname = response[0]
        score = response[1]
        detail = response[2]        
        target_row = stdrow[stdname] # todo: 나중에 존재 X 경우 예외처리 

        # score 업데이트
        if ws.cell(target_row, TARGET_COL).value:
            ws.cell(target_row, TARGET_COL, score + ws.cell(target_row, TARGET_COL).value)
        else:
            ws.cell(target_row, TARGET_COL, score )
        
        # detail 메모 추가 / 기존 comment 가져와야
        if ws.cell(target_row, TARGET_COL).comment:
            comment = Comment(ws.cell(target_row, TARGET_COL).comment.content + '\n' + stdname + ":" +detail + " " + str(score), AUTHOR)
        else:
            comment = Comment(stdname + ":" +detail + " " + str(score), AUTHOR) #
        ws.cell(target_row, TARGET_COL).comment = comment
        
        # console print
        print(target_row, TARGET_COL, stdname, score, detail)
                

    elif "ㅁㅁ" in line:
        """
        << 여러명 상벌점(ㅁㅁ) 처리 >>
        ㅁㅁ 홍길동 김다은 김누구 +1 원어민zel상점
        ㅁㅁ 홍길동 김다은 김누구 -1 원어민zel벌점
        """
        
        response = find_command_and_return_index(line, "ㅁㅁ")
        stdnamelist = response[0]
        score = response[1]
        detail = response[2]        
        print(response)
        
        for stdname in stdnamelist:
            target_row = stdrow[stdname] # todo: 나중에 존재 X 경우 예외처리 

            # score 업데이트
            if ws.cell(target_row, TARGET_COL).value:
                ws.cell(target_row, TARGET_COL, score + ws.cell(target_row, TARGET_COL).value)
            else:
                ws.cell(target_row, TARGET_COL, score )
            
            # detail 메모 추가 / 기존 comment 가져와야
            # comment = Comment(stdname + ":" +detail + " " + str(score), AUTHOR)
            if ws.cell(target_row, TARGET_COL).comment:
                comment = Comment(ws.cell(target_row, TARGET_COL).comment.content + '\n' + stdname + ":" +detail + " " + str(score), AUTHOR)
            else:
                comment = Comment(stdname + ":" +detail + " " + str(score), AUTHOR) #
            ws.cell(target_row, TARGET_COL).comment = comment
            
            # console print
            print(target_row, TARGET_COL, stdname, score, detail)
    

    elif "ㄴㄴ" in line:
        """
        << 룸 기준 상벌점(ㄴㄴ) 처리 >>
        ㄴㄴ 608 +3 청소우수
        """
        
        response = find_command_and_return_index(line, "ㄴㄴ")
        roomnum = int(response[0]) # room number
        score = response[1]
        detail = response[2]       
        stdlist = room[roomnum] # 해당 room 학생들 리스트  

        for stdname in stdlist: 
            target_row = stdrow[stdname]

            # score 업데이트
            if ws.cell(target_row, TARGET_COL).value:
                ws.cell(target_row, TARGET_COL, score + ws.cell(target_row, TARGET_COL).value)
            else:
                ws.cell(target_row, TARGET_COL, score)
            

            # detail 메모 추가
            # comment = Comment(stdname + ":" + detail + " " + str(score), AUTHOR)
            if ws.cell(target_row, TARGET_COL).comment:
                comment = Comment(ws.cell(target_row, TARGET_COL).comment.content + '\n' + stdname + ":" +detail + " " + str(score), AUTHOR)
            else:
                comment = Comment(stdname + ":" +detail + " " + str(score), AUTHOR) #
            ws.cell(target_row, TARGET_COL).comment = comment
            
            # console print
            print(target_row, TARGET_COL, stdname, score, detail)
        


f.close()
wb.save(EXCEL_FILE_NAME)