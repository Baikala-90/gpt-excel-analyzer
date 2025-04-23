import os
from flask import Flask, request, render_template
import pandas as pd
from datetime import timedelta

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# 기준 속도
BW_PRINT_SPEED = 22320
COLOR_PRINT_SPEED = 8040
BIND_SPEED = 300
COATING_SPEED = 84

COATING_LENGTH = {
    '날개없음': 468,
    '날개있음 (B5)': 630,
    '날개있음 (A5, 46판)': 560
}

def minutes_to_time(minutes):
    td = timedelta(minutes=minutes)
    return f"{td.seconds // 3600}:{(td.seconds % 3600) // 60:02}"

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file.filename.endswith('.xlsx'):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            df = pd.read_excel(filepath)
            summary = analyze_dataframe(df)
            return render_template('index.html', result=summary)
    return render_template('index.html', result=None)

def analyze_dataframe(df):
    result = {}
    result['종수'] = len(df)
    result['부수'] = int(df['발주량'].sum())
    result['흑백페이지수'] = int(df['흑백페이지수'].sum())
    result['칼라페이지수'] = int(df['컬러페이지수'].sum())
    result['흑백인쇄시간'] = minutes_to_time(result['흑백페이지수'] / (BW_PRINT_SPEED / 60))
    result['칼라인쇄시간'] = minutes_to_time(result['칼라페이지수'] / (COLOR_PRINT_SPEED / 60))
    result['제본시간'] = minutes_to_time(result['부수'] / (BIND_SPEED / 60))

    # 코팅 시간
    glossy_len = 0
    matte_len = 0
    for _, row in df.iterrows():
        qty = row['발주량']
        wing = row['날개']
        spec = str(row['규격'])
        if wing == '있음':
            key = '날개있음 (B5)' if 'B5' in spec else '날개있음 (A5, 46판)'
        else:
            key = '날개없음'
        length = COATING_LENGTH.get(key, 468) * qty
        if row['코팅'] == '유광':
            glossy_len += length
        elif row['코팅'] == '무광':
            matte_len += length
    result['유광코팅시간'] = minutes_to_time((glossy_len / COATING_SPEED) / 60)
    result['무광코팅시간'] = minutes_to_time((matte_len / COATING_SPEED) / 60)

    # 1시간 이상 도서
    over_bw, over_color = [], []
    for _, row in df.iterrows():
        bw_time = row['흑백페이지수'] / 11160
        color_time = row['컬러페이지수'] / 4020
        if bw_time >= 1:
            over_bw.append((row['도서명'], row.get('파일명(그룹명)', ''), int(row['흑백페이지수']), minutes_to_time(bw_time * 60)))
        if color_time >= 1:
            over_color.append((row['도서명'], row.get('파일명(그룹명)', ''), int(row['컬러페이지수']), minutes_to_time(color_time * 60)))
    result['1시간흑백'] = over_bw
    result['1시간칼라'] = over_color

    # 폴더별 내지 파일 수
    df['내지파일개수'] = df['공정구분'].apply(lambda x: 2 if x == '혼합' else 1)
    result['폴더별'] = df.groupby('폴더')['내지파일개수'].sum().to_dict()
    result['표지수'] = result['종수']
    result['내지수'] = df['내지파일개수'].sum()
    result['총합'] = result['표지수'] + result['내지수']
    return result

if __name__ == '__main__':
    app.run(debug=True)
