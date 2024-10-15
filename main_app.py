import streamlit as st
from datetime import datetime
import time
import threading

# 각 보고서 유형에 대한 함수를 import합니다
from report.customer_balance_sheet import generate_customer_balance_sheet_report
from report.customer_income_statement import generate_customer_income_statement_report
from report.customer_cash_flow import generate_customer_cash_flow_report
from report.customer_equity_change import generate_customer_equity_change_report
from report.strategic_recommendations import generate_strategic_recommendations_report

# CSS 스타일 정의
def local_css(file_name):
    with open(file_name, "r") as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

# 리포트 스타일링을 위한 HTML 템플릿
report_template = """
<div class="report-container">
    <h1 class="report-title">{title}</h1>
    <div class="report-date">분석 일자: {date}</div>
    <div class="report-content">
        {content}
    </div>
</div>
"""

def update_progress(my_bar, stop_event):
    progress_text = "AI 리포트 작성 중... 잠시만 기다려주세요."
    for percent_complete in range(200):
        if stop_event.is_set():
            break
        time.sleep(0.1)
        my_bar.progress((percent_complete + 1) / 200, text=progress_text)

# 로고 이미지 표시 함수
def display_logo():
    col1, col2, col3 = st.columns([1,2,1])
    with col1:
        st.image("logo.png", width=200)
    with col2:
        st.write("")
    with col3:
        st.write("")

# 각 지표별 설명
report_descriptions = {
    "고객상태표(Customer Balance Sheet)": "기업이 보유한 총 고객을 자본고객과 부채고객으로 분류하여 고객의 구성 상태를 파악합니다.",
    "고객손익계산서(Customer Income Statement)": "자본고객과 부채고객에 대한 손익 계산을 통해 고객 유형별 공헌이익 및 영업이익을 비교합니다.",
    "고객흐름표(Statement of Customer Flows)": "고객을 획득할 수 있는 점포유형을 대형점포, 중형점포, 소형점포로 구분하고 점포규모별 고객의 흐름을 파악합니다.",
    "고객변동표(Statement of Changes in Customer)": "자본고객을 5분위로 구분하여 각 분위 계층의 현황과 변동사항을 파악합니다.",
    "전략적 제언(Strategic recommendation)": "고객제표 분석 결과에 기반한 전략적 제언을 제공합니다."
}

# 로그인 상태를 저장할 변수
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

# 로그인 함수
def login():
    display_logo()
    st.title("로그인")
    
    st.write("회사에서 제공받은 계정을 입력해주세요.")
    
    아이디 = st.text_input("아이디")
    비밀번호 = st.text_input("비밀번호", type="password")
    
    if st.button("로그인"):
        if 아이디 == "1" and 비밀번호 == "1":
            st.session_state['logged_in'] = True
            st.rerun()  # 여기를 수정했습니다
        else:
            st.error("아이디 또는 비밀번호가 잘못되었습니다.")

# 홈 화면 함수
def home_page():
    display_logo()
    st.title(":bar_chart: AI 고객제표 분석")
    
    st.markdown("---")
    
    with st.expander("**:clipboard: 고객제표 분석 방법**"):
        for report_type, description in report_descriptions.items():
            st.write(f"- **{report_type}**: {description}")
    
    selected_part = st.selectbox(
        "분석할 지표를 선택하세요. *",
        list(report_descriptions.keys())
    )
    
    if selected_part:
        upload_page(selected_part)

# 업로드 및 프롬프트 작성 페이지 함수
def upload_page(part_name):
    st.write(f"선택된 분석: **{part_name}**")
    st.write(report_descriptions[part_name])
    
    # 파일 업로드
    uploaded_file = st.file_uploader(f"{part_name} 관련 데이터 파일(xlsx)을 업로드하세요. *")
    
    # 추가 프롬프트 입력
    user_prompt = st.text_area("추가적인 분석 요청사항이 있다면 입력해주세요.", 
        placeholder="예: 특정 고객 그룹에 대한 더 자세한 분석이 필요합니다.", 
        height=100
    )

    # AI 리포트 실행 버튼
    if st.button("AI 리포트 작성"):
        if uploaded_file:
            progress_text = "AI 리포트 작성 중... 잠시만 기다려주세요."
            progress_bar = st.progress(0)
            status_text = st.empty()

            # 분석 결과를 저장할 변수
            analysis_result = None

            # 분석 함수를 별도의 스레드에서 실행
            def run_analysis():
                nonlocal analysis_result
                analysis_result = generate_analysis(uploaded_file, part_name, user_prompt)

            analysis_thread = threading.Thread(target=run_analysis)
            analysis_thread.start()

            # Progress bar 업데이트
            for percent_complete in range(100):
                time.sleep(0.55)  # 0.1초마다 업데이트 (총 10초)
                progress_bar.progress(percent_complete + 1)
                status_text.text(f"{progress_text} {percent_complete + 1}%")

            # 분석 스레드가 완료될 때까지 대기
            analysis_thread.join()

            progress_bar.empty()
            status_text.empty()

            st.success(f"{part_name} 분석이 완료되었습니다!")
            
            # 리포트 표시
            report_content = report_template.format(
                title=f"{part_name} 분석 리포트",
                date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                content=analysis_result.replace('\n', '<br>')
            )
            st.markdown(report_content, unsafe_allow_html=True)
        else:
            st.warning("파일을 업로드해주세요.")

# 분석 결과를 통합하여 전략적 제언을 생성하는 함수
def generate_analysis(file, report_type, user_prompt):
    try:
        # 보고서 유형에 따라 적절한 보고서 생성 함수 호출
        if report_type == "고객상태표(Customer Balance Sheet)":
            return generate_customer_balance_sheet_report(file, user_prompt)
        elif report_type == "고객손익계산서(Customer Income Statement)":
            return generate_customer_income_statement_report(file, user_prompt)
        elif report_type == "고객흐름표(Statement of Customer Flows)":
            return generate_customer_cash_flow_report(file, user_prompt)
        elif report_type == "고객변동표(Statement of Changes in Customer)":
            return generate_customer_equity_change_report(file, user_prompt)
        elif report_type == "전략적 제언(Strategic recommendation)":
            # 각 보고서를 생성하고 이를 바탕으로 전략적 제언을 생성합니다.
            balance_sheet_result = generate_customer_balance_sheet_report(file, user_prompt)
            income_statement_result = generate_customer_income_statement_report(file, user_prompt)
            cash_flow_result = generate_customer_cash_flow_report(file, user_prompt)
            equity_change_result = generate_customer_equity_change_report(file, user_prompt)
            return generate_strategic_recommendations_report(balance_sheet_result, income_statement_result, cash_flow_result, equity_change_result, user_prompt)
        else:
            raise ValueError("Invalid report type")
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        return f"분석 생성 중 오류가 발생했습니다: {str(e)}\n\n상세 오류:\n{error_details}"

# 과거 내역 페이지 함수
def history_page():
    display_logo()
    st.title(":file_folder: 과거 분석 내역")
    
    st.markdown("---")
    
    st.write("Sorry, No data found :sweat_smile:")  # 실제 데이터를 불러오는 방식으로 변경 가능
    
# 고객제표 시각화 페이지
def customer_statements_visualization():
    st.title("고객제표 시각화")
    st.image("dash.png", use_column_width=True)
    st.write("이 페이지는 고객제표의 시각화를 보여줍니다.")
    st.write("추가적인 설명이나 분석을 여기에 포함할 수 있습니다.")

# Streamlit 시작
if __name__ == "__main__":
    st.set_page_config(layout="wide", page_title="AI 고객제표 분석")
    local_css("style.css")  # CSS 파일 로드
    
    if 'page' not in st.session_state:
        st.session_state['page'] = 'home'
    
    if st.session_state['logged_in']:
        st.sidebar.title(":pushpin: 메뉴")
        selection = st.sidebar.radio("접속할 탭을 선택하세요.", ["홈 화면", "과거 분석 내역"])
        
        if selection == "홈 화면":
            st.session_state['page'] = 'home'
        elif selection == "과거 분석 내역":
            st.session_state['page'] = 'history'

        st.sidebar.markdown("---")
        st.sidebar.markdown("**:bar_chart: 고객제표 시각화**")
        
        # 고객제표 시각화 버튼
        if st.sidebar.button("고객제표 시각화 보기"):
            st.session_state['page'] = 'visualization'
        
        # 페이지 렌더링
        if st.session_state['page'] == 'home':
            home_page()
        elif st.session_state['page'] == 'history':
            history_page()
        elif st.session_state['page'] == 'visualization':
            customer_statements_visualization()
        
    else:
        login()
