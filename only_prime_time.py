import os
import pandas as pd

# 현재 폴더에 있는 모든 파일들의 목록을 읽음
file_list = os.listdir()

# xlsx 확장자를 가지는 파일만 필터링
xlsx_files = [file for file in file_list if file.endswith('.xlsx')]

# 필터링된 데이터프레임을 저장할 리스트
filtered_dataframes = []

# xlsx 파일을 pandas DataFrame으로 읽으며 상위 5행을 건너뜀
for file in xlsx_files:
    # 파일 읽기
    df = pd.read_excel(file, sheet_name="프로그램일자별리스트", skiprows=5)
    
    # 시작시간과 종료시간을 문자열로 변환
    df['시작시간'] = df['시작시간'].astype(str)
    df['종료시간'] = df['종료시간'].astype(str)

    # ':'을 기준으로 분리하여 시간 부분만 추출 후 int 형식으로 변환 / 불가한 경우 NaN으로 변환
    df['시작시간hour'] = pd.to_numeric(df['시작시간'].str.split(':').str[0], errors='coerce')
    df['종료시간hour'] = pd.to_numeric(df['종료시간'].str.split(':').str[0], errors='coerce')

    # 시작시간 ~ 종료시간 필터링
    start_time = 19
    end_time = 24
    filtered_df = df[(df['시작시간hour'] >= start_time) & (df['종료시간hour'] < end_time)]


    # print(filtered_df)
    # 필터링된 데이터프레임 리스트에 추가
    filtered_dataframes.append(filtered_df)

# 모든 데이터프레임을 하나로 합침
if filtered_dataframes:
    combined_df = pd.concat(filtered_dataframes, ignore_index=True)
else:
    combined_df = pd.DataFrame()  # 필터링된 데이터가 없을 경우 빈 데이터프레임 생성

# output 폴더가 없으면 생성
output_dir = 'output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# combined_df를 output 폴더에 result.csv로 저장
output_path = os.path.join(output_dir, 'time.csv')
combined_df.to_csv(output_path, index=False, encoding='utf-8-sig')

print(f"Filtered data saved to {output_path}")