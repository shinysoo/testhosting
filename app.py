from flask import Flask, render_template, request, redirect
import pandas as pd
import datetime
import os

aaa=pd.read_excel("C:/Users/User/Desktop/복호화폴더/인슈플러스_판매실적.xlsx")
aaa.rename(columns={'Unnamed: 1':'구분값'},inplace=True)
aaa['구분값']=pd.to_datetime(aaa['구분값'])
aaa['보험시작일']=pd.to_datetime(aaa['보험시작일'])
aaa['보험종료일']=pd.to_datetime(aaa['보험종료일'])
aaa.drop(['NO'],axis=1,inplace=True)

app = Flask(__name__, template_folder='C:/Users/User/Desktop/인슈플러스')
app.config['UPLOAD_FOLDER'] = 'C:/Users/User/Desktop/인슈플러스/uploads'  # 파일 업로드를 위한 절대 경로 설정

# 해당 경로가 존재하지 않는 경우 생성
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if file:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)  # 정확한 파일 경로 생성
            file.save(filepath)  # 파일 저장
            process_file(filepath)  # 데이터 처리 함수 호출
            return redirect('/')  # 처리 후 메인 페이지로 리다이렉션
    return render_template('index.html')

def process_file(filepath):
    aaa = pd.read_excel(filepath)
    date_range = pd.date_range(start='2019-09-17', end='2024-04-30')
    channels = ['온라인', 'B2B', '제휴']
    periods = ['단기', '장기']
    final_dfs = []

    for date in date_range:
        data_row = {'구분값': date}
        for channel in channels:
            for period in periods:
                subscrip = aaa[(aaa['채널구분'] == channel) &
                               (aaa['가입기간'] == period) &
                               (aaa['구분값'] == date)]['인원수'].sum()
                active_people = aaa[(aaa['채널구분'] == channel) &
                                    (aaa['가입기간'] == period) &
                                    (aaa['보험시작일'] <= date) &
                                    (date <= aaa['보험종료일'])]['인원수'].sum()
                data_row[f'{channel}_{period}_가입'] = subscrip
                data_row[f'{channel}_{period}_Active'] = active_people

        final_dfs.append(pd.DataFrame([data_row]))

    final_df = pd.concat(final_dfs)
    final_df.set_index('구분값', inplace=True)
    final_df.fillna(0, inplace=True)
    final_df.to_excel('C:/Users/User/Desktop/인슈플러스/실적데이터_new.xlsx', index=True)

if __name__ == '__main__':
    app.run(debug=True)
