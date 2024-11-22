import pandas as pd

week = "9차"
path0 = r"E:\2024_1_new\_NTree\2024-1학기 창의 NTree 공동문서.xlsx"
path1 = r"E:\2024_1_new\_NTree\2-2_[과제1]_아두이노_사전_과제(Thinkercad)_과제제출상태.xlsx"
path2 = r"E:\2024_1_new\_NTree\3_[과제2]_제품_아이디어_제출_과제제출상태.xlsx"
path3 = r"E:\2024_1_new\_NTree\3_[과제3]_개발계획서_제출_과제제출상태.xlsx"
path4 = r"E:\2024_1_new\_NTree\1._[과제4]_최종보고서_제출(팀과제)_과제제출상태.xlsx"
path5 = r"E:\2024_1_new\_NTree\2024-1학기 9차 창의 NTree 캠프 (스마트보안전공, 스마트시티학과, 컴퓨터공학전공_B)_팀플평가.xlsx"
work1_scores = [10, 8, 0]  # 정상제출, 기간 후 제출
work2_scores = [10, 8, 0]
work3_scores = [24, 20, 0]
work4_scores = [5, 5, 0]
group_column_name = '그룹\n(팀)'  # '그룹\n(팀/불참)'

summary_excel = pd.read_excel(path0, sheet_name=week, header=1, )

# 과제1_틴거캐드 제출 처리
print("1. 과제 1(틴거캐드) 성적 처리중...", end="")
submission1_excel = pd.read_excel(path1, header=2, usecols=[1, 4])
for index, row in submission1_excel.iterrows():
    if row['제출상태'] == '정상제출':
        student_id = row['학번']
        student_row = summary_excel.loc[summary_excel['학번'] == student_id]
        if not student_row.empty:
            summary_excel.at[student_row.index[0], '과제1\n(팅커캐드)'] = work1_scores[0]
    elif row['제출상태'] == '기간 후 제출':
        student_id = row['학번']
        student_row = summary_excel.loc[summary_excel['학번'] == student_id]
        if not student_row.empty:
            summary_excel.at[student_row.index[0], '과제1\n(팅커캐드)'] = work1_scores[1]
for index, row in summary_excel.iterrows():
    if row['과제1\n(팅커캐드)'] != work1_scores[0] and row['과제1\n(팅커캐드)'] != work1_scores[1]:
        summary_excel.at[index, '과제1\n(팅커캐드)'] = work1_scores[2]
print("처리완료")

# 과제2_아이디어 제출 처리
print("2. 과제 2(아이디어) 성적 처리중...", end="")
submission2_excel = pd.read_excel(path2, header=2, usecols=[1, 4])
for index, row in submission2_excel.iterrows():
    if row['제출상태'] == '정상제출':
        student_id = row['학번']
        student_row = summary_excel.loc[summary_excel['학번'] == student_id]
        if not student_row.empty:
            summary_excel.at[student_row.index[0], '과제2\n(아이디어)'] = work2_scores[0]
    elif row['제출상태'] == '기간 후 제출':
        student_id = row['학번']
        student_row = summary_excel.loc[summary_excel['학번'] == student_id]
        if not student_row.empty:
            summary_excel.at[student_row.index[0], '과제2\n(아이디어)'] = work2_scores[1]
for index, row in summary_excel.iterrows():
    if row['과제2\n(아이디어)'] != work2_scores[0] and row['과제2\n(아이디어)'] != work2_scores[1]:
        summary_excel.at[index, '과제2\n(아이디어)'] = work2_scores[2]
print("처리완료")

# 과제3_개발계획서 처리
print("3. 과제 3(개발계획서) 성적 처리중...", end="")
submission3_excel = pd.read_excel(path3, header=2, usecols=[1, 4])
for index, row in submission3_excel.iterrows():
    if row['제출상태'] == '정상제출':
        student_id = row['학번']
        student_row = summary_excel.loc[summary_excel['학번'] == student_id]
        if not student_row.empty:
            summary_excel.at[student_row.index[0], '과제3\n(개발계획)'] = work3_scores[0]
    elif row['제출상태'] == '기간 후 제출':
        student_id = row['학번']
        student_row = summary_excel.loc[summary_excel['학번'] == student_id]
        if not student_row.empty:
            summary_excel.at[student_row.index[0], '과제3\n(개발계획)'] = work3_scores[1]
for index, row in summary_excel.iterrows():
    if row['과제3\n(개발계획)'] == work3_scores[0] or row['과제3\n(개발계획)'] == work3_scores[1]:
        group_num = row[group_column_name]
        score = row['과제3\n(개발계획)']
        for index2, row2 in summary_excel.iterrows():
            if row2[group_column_name] == group_num:
                summary_excel.at[index2, '과제3\n(개발계획)'] = score
for index, row in summary_excel.iterrows():
    if row['과제3\n(개발계획)'] != work3_scores[0] and row['과제3\n(개발계획)'] != work3_scores[1]:
        summary_excel.at[index, '과제3\n(개발계획)'] = work3_scores[2]
print("처리완료")

# 과제4_CF영상&최종발표 동시 처리
print("4. 과제 4(CF영상&최종발표) 성적 처리중...", end="")
submission4_excel = pd.read_excel(path4, header=2, usecols=[1, 4])
for index, row in submission4_excel.iterrows():
    if row['제출상태'] == '정상제출':
        student_id = row['학번']
        student_row = summary_excel.loc[summary_excel['학번'] == student_id]
        if not student_row.empty:
            summary_excel.at[student_row.index[0], '과제4\n(CF 영상)'] = work4_scores[0]
            summary_excel.at[student_row.index[0], '과제4\n(최종발표)'] = work4_scores[0]
    elif row['제출상태'] == '기간 후 제출':
        student_id = row['학번']
        student_row = summary_excel.loc[summary_excel['학번'] == student_id]
        if not student_row.empty:
            summary_excel.at[student_row.index[0], '과제4\n(CF 영상)'] = work4_scores[1]
            summary_excel.at[student_row.index[0], '과제4\n(최종발표)'] = work4_scores[1]
for index, row in summary_excel.iterrows():
    if row['과제4\n(CF 영상)'] == work4_scores[0] or row['과제4\n(CF 영상)'] == work4_scores[1]:
        group_num = row[group_column_name]
        score = row['과제4\n(CF 영상)']
        for index2, row2 in summary_excel.iterrows():
            if row2[group_column_name] == group_num:
                summary_excel.at[index2, '과제4\n(CF 영상)'] = score
                summary_excel.at[index2, '과제4\n(최종발표)'] = score
for index, row in summary_excel.iterrows():
    if row['과제4\n(CF 영상)'] != work4_scores[0] and row['과제4\n(CF 영상)'] != work4_scores[1]:
        summary_excel.at[index, '과제4\n(CF 영상)'] = work4_scores[2]
        summary_excel.at[index, '과제4\n(최종발표)'] = work4_scores[2]
print("처리완료")

# 동료평가 처리

print("5. 동료평가 성적 전처리중...", end="")
peer_eval_excel = pd.read_excel(path5, usecols=[2, 4, 11])
review_list = []
for index, row in peer_eval_excel.iterrows():
    student_id1 = row['피평가자 학번']
    student_id2 = row['평가자 학번']
    group1 = ''
    group2 = ''
    for index2, row2 in summary_excel.iterrows():
        if row2['학번'] == student_id1:
            group1 = row2[group_column_name]
        elif row2['학번'] == student_id2:
            group2 = row2[group_column_name]
    if group1 != group2:
        review_list.append(student_id1)
        peer_eval_excel = peer_eval_excel.drop([index])

print("동료평가 성적 처리중...", end="")
for index, row in summary_excel.iterrows():
    student_id = row['학번']
    count = 0
    sum_grade = 0.0
    avg_grade = 0.0
    student2_list = []
    for index2, row2 in peer_eval_excel.iterrows():
        if row2['피평가자 학번'] == student_id:
            count = count + 1
            sum_grade = sum_grade + int(row2['총점 / 만점'].split('/')[0])

    if count != 0:
        avg_grade = sum_grade / count
        avg_grade = avg_grade * 100 // 1 / 100
    if count == 0:
        avg_grade = 0.0
    summary_excel.at[index, '동료평가 - 20점\n(13점 미만 F)'] = avg_grade
print("처리완료")

# 조 오기입 의심 학생 필터링
review_list = list(set(review_list))
print("종합파일 조 오기입이 의심되어 추가확인이 필요한 학생 수: ", len(review_list))
i = 1
for id in review_list:
    print(id, end=", ")
    if (i % 5 == 0):
        print()
    i = i + 1

#결과출력
summary_excel.to_excel("result.xlsx", index=False)
