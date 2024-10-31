import openpyxl
import pandas as pd
import random

lucky_week_ball = []

# read excel
def read_excel_file(file_path, sheet_name):
    return pd.read_excel(file_path, sheet_name=sheet_name)


# 엑셀에서 지정된 범위의 데이터를 추출하는 함수
def get_data_in_range(df, start_cell, end_cell):
    start_row, start_col = start_cell
    end_row, end_col = end_cell
    return df.iloc[start_row-1:end_row, start_col-1:end_col]


# 숫자 빈도수를 계산하는 함수
def calculate_frequency(data):
    flat_data = data.values.flatten()
    # 숫자만 필터링
    numbers = [x for x in flat_data if isinstance(x, (int, float))]
    return pd.Series(numbers).value_counts()


# 이상값 제거 함수 (IQR 방법)
def remove_outliers_iqr(frequency_dict):
    values = list(frequency_dict.values())
    q1 = pd.Series(values).quantile(0.25)
    q3 = pd.Series(values).quantile(0.75)
    iqr = q3 - q1
    lower_bound = q1 - 1.5 * iqr
    upper_bound = q3 + 1.5 * iqr
    # IQR 범위 내의 값만 남김
    cleaned_dict = {k: v for k, v in frequency_dict.items() if lower_bound <= v <= upper_bound}
    outliers = {k: v for k, v in frequency_dict.items() if not lower_bound <= v <= upper_bound} #tempo
    return cleaned_dict, outliers

# 15주차 가중 뽑기
def week15_choice(week15_ball):
   return random.choice(week15_ball)

#top3 frequent_ball
def find_top_3_frequent_numbers(frequency_dict):
    sorted_items = sorted(frequency_dict.items(), key=lambda x: x[1], reverse=True)
    return [item[0] for item in sorted_items[:3]]


#main
def calculate():

    #구간별 가중치 변수 * 6/13일자 기준
    a_1 = 11 # 1 ~ 5
    a_2 = 14 # 6 ~ 10
    b_1 = 15 # 11 ~ 15
    b_2 = 13 # 16 ~ 20
    c_1 = 11 # 21 ~ 25
    c_2 = 16 # 26 ~ 30
    d_1 = 12 # 31 ~ 35
    d_2 = 8 # 36 ~ 40
    e_1 = 5 # 41 ~ 45

    #list
    lucky_ball = [] #전체 번호 배열
    week15_ball = [32,39,43] #15주차 미출현 번호 배열  
    result_ball = [] #결과 배열
    top3_ball = [] #최빈값 상위 3개 배열

    
    #lucky_ball = 1~45
    for i in range(1,46,1):
        lucky_ball.append(i) 

    #15주차 미 출현 공 중 1개를 뽑고 최우선 결과값 result에 저장
    chosen_week15_ball = week15_choice(week15_ball)    
    result_ball.append(chosen_week15_ball)

    
    #15주차 미 출현 공들에 대하여 가중배열에서 삭제처리
    for ball in week15_ball:
        if ball in lucky_ball:
            lucky_ball.remove(ball)

    
    #excel data관련 지정
    file_path = "C:\\로또 business\\lottto_ver1.xlsx"
    sheet_name= 'excel'
    start_cell = (4,3) # 시작 셀 C4 
    end_cell = (546, 9)  # 끝 셀 I534

    #call read excel func
    df = read_excel_file(file_path, sheet_name)

    #call read excel range func
    data_in_range = get_data_in_range(df, start_cell, end_cell)

    #call calculate frequency func
    frequency = calculate_frequency(data_in_range)
    

    #frequency result return to dict
    frequency_dict = frequency.to_dict()
    
    # 이상값 찾기
    cleaned_frequency_dict, outliers = remove_outliers_iqr(frequency_dict)


    # 이상값의 키를 lucky_ball에서 제거
    for key in outliers.keys():
        if key in lucky_ball:
            lucky_ball.remove(int(key))

    # 최빈값 상위 3개 찾기
    top_3_frequent_numbers = find_top_3_frequent_numbers(cleaned_frequency_dict)
    
    #최빈값 상위 3개에 대하여 가중 뽑기 배열에서 제거
    for number in top_3_frequent_numbers:
        top3_ball.append(number)
        if number in lucky_ball:
            lucky_ball.remove(number)

    
    #최빈값 상위 3개 중 1개 공 뽑기
    result_ball.append(random.choice(top3_ball))
    
    #구간별 출현횟수에 대한 가중치 

    # 1 ~ 5 -> 4
    for i in range(1,6):
        if i in lucky_ball:
            lucky_ball.extend([i] * a_1)

    # 6 ~ 10 -> 4
    for i in range(6,11):
        if i in lucky_ball:
            lucky_ball.extend([i] * a_2)

    # 11 ~ 15 -> 3
    for i in range(11,16):
        if i in lucky_ball:
            lucky_ball.extend([i] * b_1)

    # 16 ~ 20 -> 3
    for i in range(16,21):
        if i in lucky_ball:
            lucky_ball.extend([i] * b_2)

    # 21 ~ 25 -> 5
    for i in range(21,26):
        if i in lucky_ball:
            lucky_ball.extend([i] * c_1)

    # 26 ~ 30 -> 3
    for i in range(26,31):
        if i in lucky_ball:
            lucky_ball.extend([i] * c_2)

    # 31 ~ 35 -> 8
    for i in range(31, 36):
        if i in lucky_ball:
            lucky_ball.extend([i] * d_1)

    # 36 ~ 40 -> 2
    for i in range(36, 41):
        if i in lucky_ball:
            lucky_ball.extend([i] * d_2)

    # 41 ~ 45 -> 3
    for i in range(41, 46):
        if i in lucky_ball:
            lucky_ball.extend([i] * e_1)

    
    #가중치를 이용하여 남은 4개의 공 랜덤 뽑기
    for i in range(4):
        chosen_ball = random.choice(lucky_ball)
        result_ball.append(chosen_ball)

        #slicing을 이용하여 lucky_ball list에서 지정한 값이 아닌 숫자를 제외하고 가져옴
        lucky_ball[:] = (value for value in lucky_ball if value != chosen_ball)
    

    #최종 출력
    result_ball.sort()
    result_ball = list(map(int, result_ball))

    global lucky_week_ball
    lucky_week_ball.extend(result_ball)
    
    return result_ball

def main():
    count = 1144 #6/14일자 기준 다음(예측) 회차
    global lucky_week_ball

    for i in range(5):
        result_ball = calculate()
        print("{}회차 {}번째 행운의 조합 : {}".format(count, i+1,result_ball))

    #lucky_week_ball = set(lucky_week_ball)
    lucky_week_ball = list(set(lucky_week_ball))
    lucky_week_ball.sort()

    print("\n이번주 행운의 번호", lucky_week_ball)

        
    
    
if __name__ == "__main__":
    main()
