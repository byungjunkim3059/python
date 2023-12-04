from datetime import datetime, timedelta

def get_week_range(year, week):
    # 해당 년도와 주차로 시작일을 계산
    start_of_week = datetime.fromisocalendar(year, week, 1)
    
    # 해당 주의 마지막 날을 계산
    end_of_week = start_of_week + timedelta(days=6)

    return start_of_week, end_of_week

# 예제: 2023년의 10주차 시작일과 끝일 확인
year = 2023
week = 10

start_date, end_date = get_week_range(year, week)
print(f"{year}년 {week}주차의 시작일은 {start_date.strftime('%Y-%m-%d')}이고, 끝일은 {end_date.strftime('%Y-%m-%d')}입니다.")
