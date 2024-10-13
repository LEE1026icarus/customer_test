import openai
import pandas as pd
import io
import os
from dotenv import load_dotenv
load_dotenv()

def read_excel(file):
    try:
        # pandas로 엑셀 파일 읽기 (openpyxl 엔진 명시적으로 지정)
        df_ls = pd.read_excel(file, sheet_name='LS', engine='openpyxl')
        return df_ls
    except ValueError as e:
        print(f"파일 형식을 확인할 수 없습니다. 오류: {e}")
    except FileNotFoundError as e:
        print(f"파일을 찾을 수 없습니다. 오류: {e}")
    except Exception as e:
        print(f"엑셀 파일을 읽는 중 오류가 발생했습니다. 오류: {e}")
        
    return None


def process_data(df_ls):

    report_data = {
    "company_name": df_ls[df_ls['지표'] == '기업명']['당기'].values[0],
    "fifth_quintile_customers": int(df_ls[df_ls['지표'] == '5분위 고객 수']['값'].values[0]),
    "fourth_quintile_customers": int(df_ls[df_ls['지표'] == '4분위 고객 수']['값'].values[0]),
    "third_quintile_customers": int(df_ls[df_ls['지표'] == '3분위 고객 수']['값'].values[0]),
    "second_quintile_customers": int(df_ls[df_ls['지표'] == '2분위 고객 수']['값'].values[0]),
    "first_quintile_customers": int(df_ls[df_ls['지표'] == '1분위 고객 수']['값'].values[0]),
    "capital_customers": int(df_ls[df_ls['지표'] == '자본고객 수']['값'].values[0]),
    "fifth_quintile_sales": float(df_ls[df_ls['지표'] == '5분위 매출액']['값'].values[0]),
    "fourth_quintile_sales": float(df_ls[df_ls['지표'] == '4분위 매출액']['값'].values[0]),
    "third_quintile_sales": float(df_ls[df_ls['지표'] == '3분위 매출액']['값'].values[0]),
    "second_quintile_sales": float(df_ls[df_ls['지표'] == '2분위 매출액']['값'].values[0]),
    "first_quintile_sales": float(df_ls[df_ls['지표'] == '1분위 매출액']['값'].values[0]),
    "capital_sales": float(df_ls[df_ls['지표'] == '자본고객 매출액']['값'].values[0]),
    "fifth_quintile_avg_sales": float(df_ls[df_ls['지표'] == '5분위 1인 평균 매출액']['값'].values[0]),
    "fourth_quintile_avg_sales": float(df_ls[df_ls['지표'] == '4분위 1인 평균 매출액']['값'].values[0]),
    "third_quintile_avg_sales": float(df_ls[df_ls['지표'] == '3분위 1인 평균 매출액']['값'].values[0]),
    "second_quintile_avg_sales": float(df_ls[df_ls['지표'] == '2분위 1인 평균 매출액']['값'].values[0]),
    "first_quintile_avg_sales": float(df_ls[df_ls['지표'] == '1분위 1인 평균 매출액']['값'].values[0]),
    "capital_avg_sales": float(df_ls[df_ls['지표'] == '자본고객 1인 평균 매출액']['값'].values[0]),
    "fifth_quintile_sales_ratio": float(df_ls[df_ls['지표'] == '5분위 매출비율']['값'].values[0]),
    "fourth_quintile_sales_ratio": float(df_ls[df_ls['지표'] == '4분위 매출비율']['값'].values[0]),
    "third_quintile_sales_ratio": float(df_ls[df_ls['지표'] == '3분위 매출비율']['값'].values[0]),
    "second_quintile_sales_ratio": float(df_ls[df_ls['지표'] == '2분위 매출비율']['값'].values[0]),
    "first_quintile_sales_ratio": float(df_ls[df_ls['지표'] == '1분위 매출비율']['값'].values[0]),
    "capital_sales_ratio": float(df_ls[df_ls['지표'] == '자본고객 매출비율']['값'].values[0]),
    "fifth_quintile_new_customers": int(df_ls[df_ls['지표'] == '5분위 신규고객 수']['값'].values[0]),
    "fourth_quintile_new_customers": int(df_ls[df_ls['지표'] == '4분위 신규고객 수']['값'].values[0]),
    "third_quintile_new_customers": int(df_ls[df_ls['지표'] == '3분위 신규고객 수']['값'].values[0]),
    "second_quintile_new_customers": int(df_ls[df_ls['지표'] == '2분위 신규고객 수']['값'].values[0]),
    "first_quintile_new_customers": int(df_ls[df_ls['지표'] == '1분위 신규고객 수']['값'].values[0]),
    "total_new_customers": int(df_ls[df_ls['지표'] == '신규고객 수']['값'].values[0]),
    "fifth_quintile_upgraded_customers": int(df_ls[df_ls['지표'] == '5분위 상승고객 수']['값'].values[0]),
    "fourth_quintile_upgraded_customers": int(df_ls[df_ls['지표'] == '4분위 상승고객 수']['값'].values[0]),
    "third_quintile_upgraded_customers": int(df_ls[df_ls['지표'] == '3분위 상승고객 수']['값'].values[0]),
    "second_quintile_upgraded_customers": int(df_ls[df_ls['지표'] == '2분위 상승고객 수']['값'].values[0]),
    "first_quintile_upgraded_customers": int(0),
    "total_upgraded_customers": int(df_ls[df_ls['지표'] == '등급상승 고객 수']['값'].values[0]),
    "fifth_quintile_retained_customers": int(df_ls[df_ls['지표'] == '5분위 유지고객 수']['값'].values[0]),
    "fourth_quintile_retained_customers": int(df_ls[df_ls['지표'] == '4분위 유지고객 수']['값'].values[0]),
    "third_quintile_retained_customers": int(df_ls[df_ls['지표'] == '3분위 유지고객 수']['값'].values[0]),
    "second_quintile_retained_customers": int(df_ls[df_ls['지표'] == '2분위 유지고객 수']['값'].values[0]),
    "first_quintile_retained_customers": int(df_ls[df_ls['지표'] == '1분위 유지고객 수']['값'].values[0]),
    "total_retained_customers": int(df_ls[df_ls['지표'] == '등급유지 고객 수']['값'].values[0]),
    "fifth_quintile_downgraded_customers": int(0),
    "fourth_quintile_downgraded_customers": int(df_ls[df_ls['지표'] == '4분위 하락고객 수']['값'].values[0]),
    "third_quintile_downgraded_customers": int(df_ls[df_ls['지표'] == '3분위 하락고객 수']['값'].values[0]),
    "second_quintile_downgraded_customers": int(df_ls[df_ls['지표'] == '2분위 하락고객 수']['값'].values[0]),
    "first_quintile_downgraded_customers": int(df_ls[df_ls['지표'] == '1분위 하락고객 수']['값'].values[0]),
    "total_downgraded_customers": int(df_ls[df_ls['지표'] == '등급하락 고객 수']['값'].values[0]),
    "fifth_quintile_lost_customers": int(df_ls[df_ls['지표'] == '5분위 이탈고객 수']['값'].values[0]),
    "fourth_quintile_lost_customers": int(df_ls[df_ls['지표'] == '4분위 이탈고객 수']['값'].values[0]),
    "third_quintile_lost_customers": int(df_ls[df_ls['지표'] == '3분위 이탈고객 수']['값'].values[0]),
    "second_quintile_lost_customers": int(df_ls[df_ls['지표'] == '2분위 이탈고객 수']['값'].values[0]),
    "first_quintile_lost_customers": int(df_ls[df_ls['지표'] == '1분위 이탈고객 수']['값'].values[0]),
    "total_lost_customers": int(df_ls[df_ls['지표'] == '이탈고객 수']['값'].values[0]),
    "new_customers_ratio": float(df_ls[df_ls['지표'] == '신규고객 비율']['값'].values[0]),
    "upgraded_customers_ratio": float(df_ls[df_ls['지표'] == '등급 상승률']['값'].values[0]),
    "retained_customers_ratio": float(df_ls[df_ls['지표'] == '등급 유지율']['값'].values[0]),
    "downgraded_customers_ratio": float(df_ls[df_ls['지표'] == '등급 하락률']['값'].values[0]),
    "lost_customers_ratio": float(df_ls[df_ls['지표'] == '고객 이탈률']['값'].values[0]),
    "maintained_customers_ratio": float(df_ls[df_ls['지표'] == '고객 유지율']['값'].values[0]),
    }
    return report_data


# def generate_prompt(data, user_prompt):
#     base_prompt = f""" ###지시사항###
#     당신의 작업은 {data['company_name']}의 2023년 고객변동표를 분석하여 고객 중심의 자본변동표를 작성하는 것입니다. 모든 60개의 데이터셋을 활용하여 각 분위 계층의 현황과 변동사항을 상세히 분석하고, 데이터에 기반한 인사이트와 실행 가능한 전략적 제언을 포함해주세요.

#     ###단계 1: 고객변동표 개요###
#     고객변동표의 의미와 중요성을 간단히 설명하세요. 다음 사항을 포함하세요:
#     - 고객가치 기준으로 그룹을 분화하여 평가하는 방식
#     - 5개의 고객등급(5분위)으로 구분한 이유
#     - 신규, 상승, 하락, 유지, 이탈 등의 변동 상황 추적 목적
#     - 자본고객 기준 설명 (전체 자본고객 수: {data['capital_customers']}명)

#     ###단계 2: 전체 고객 변동 분석###
#     다음 데이터를 활용하여 전체적인 고객 변동 상황을 분석하세요:
#     - 신규 유입 고객 수: {data['total_new_customers']}명 (전체 자본고객의 {data['total_new_customers'] / data['capital_customers'] * 100:.1f}%)
#     - 이탈 고객 수: {data['total_lost_customers']}명 (전체 자본고객의 {data['total_lost_customers'] / data['capital_customers'] * 100:.1f}%)
#     - 등급 상승 고객 수: {data['total_upgraded_customers']}명 (전체 자본고객의 {data['total_upgraded_customers'] / data['capital_customers'] * 100:.1f}%)
#     - 등급 하락 고객 수: {data['total_downgraded_customers']}명 (전체 자본고객의 {data['total_downgraded_customers'] / data['capital_customers'] * 100:.1f}%)
#     - 등급 유지 고객 수: {data['total_retained_customers']}명 (전체 자본고객의 {data['total_retained_customers'] / data['capital_customers'] * 100:.1f}%)

#     전체 고객 변동 비율도 함께 분석하세요:
#     - 신규 고객 비율: {data['new_customers_ratio']:.1f}%
#     - 등급 상승률: {data['upgraded_customers_ratio']:.1f}%
#     - 등급 유지율: {data['retained_customers_ratio']:.1f}%
#     - 등급 하락률: {data['downgraded_customers_ratio']:.1f}%
#     - 고객 이탈률: {data['lost_customers_ratio']:.1f}%
#     - 고객 유지율: {data['maintained_customers_ratio']:.1f}%

#     이 데이터를 바탕으로 {data['company_name']}의 고객 기반 안정성과 성장 가능성을 평가하세요.

#     ###단계 3: 분위별 고객 현황 및 매출 분석###
#     각 분위별 고객 수, 매출 현황, 1인당 평균 매출액을 분석하세요. 특히 상위 20% 고객의 매출 기여도와 하위 고객층의 특성에 주목하세요.

#     분위별 고객 수:
#     - 5분위: {data['fifth_quintile_customers']}명 (전체의 {data['fifth_quintile_customers'] / data['capital_customers'] * 100:.1f}%)
#     - 4분위: {data['fourth_quintile_customers']}명 (전체의 {data['fourth_quintile_customers'] / data['capital_customers'] * 100:.1f}%)
#     - 3분위: {data['third_quintile_customers']}명 (전체의 {data['third_quintile_customers'] / data['capital_customers'] * 100:.1f}%)
#     - 2분위: {data['second_quintile_customers']}명 (전체의 {data['second_quintile_customers'] / data['capital_customers'] * 100:.1f}%)
#     - 1분위: {data['first_quintile_customers']}명 (전체의 {data['first_quintile_customers'] / data['capital_customers'] * 100:.1f}%)

#     분위별 매출 현황:
#     - 5분위 매출액: {data['fifth_quintile_sales']:,.0f}원 (전체 매출의 {data['fifth_quintile_sales_ratio']:.1f}%)
#     - 4분위 매출액: {data['fourth_quintile_sales']:,.0f}원 (전체 매출의 {data['fourth_quintile_sales_ratio']:.1f}%)
#     - 3분위 매출액: {data['third_quintile_sales']:,.0f}원 (전체 매출의 {data['third_quintile_sales_ratio']:.1f}%)
#     - 2분위 매출액: {data['second_quintile_sales']:,.0f}원 (전체 매출의 {data['second_quintile_sales_ratio']:.1f}%)
#     - 1분위 매출액: {data['first_quintile_sales']:,.0f}원 (전체 매출의 {data['first_quintile_sales_ratio']:.1f}%)

#     분위별 1인당 평균 매출액:
#     - 5분위: {data['fifth_quintile_avg_sales']:,.0f}원
#     - 4분위: {data['fourth_quintile_avg_sales']:,.0f}원
#     - 3분위: {data['third_quintile_avg_sales']:,.0f}원
#     - 2분위: {data['second_quintile_avg_sales']:,.0f}원
#     - 1분위: {data['first_quintile_avg_sales']:,.0f}원

#     ###단계 4: 분위별 고객 변동 심층 분석###
#     각 분위별 신규, 상승, 유지, 하락, 이탈 고객 수를 분석하고, 다음 사항에 특히 주목하세요:
#     - 상위등급(5분위, 4분위) 고객의 하락 및 이탈 패턴
#     - 1분위 고객의 높은 이탈률과 그 원인
#     - 각 분위 간 고객 이동의 주요 특징

#     5분위:
#     - 신규: {data['fifth_quintile_new_customers']}명
#     - 상승: {data['fifth_quintile_upgraded_customers']}명
#     - 유지: {data['fifth_quintile_retained_customers']}명
#     - 하락: {data['fifth_quintile_downgraded_customers']}명
#     - 이탈: {data['fifth_quintile_lost_customers']}명

#     4분위:
#     - 신규: {data['fourth_quintile_new_customers']}명
#     - 상승: {data['fourth_quintile_upgraded_customers']}명
#     - 유지: {data['fourth_quintile_retained_customers']}명
#     - 하락: {data['fourth_quintile_downgraded_customers']}명
#     - 이탈: {data['fourth_quintile_lost_customers']}명

#     3분위:
#     - 신규: {data['third_quintile_new_customers']}명
#     - 상승: {data['third_quintile_upgraded_customers']}명
#     - 유지: {data['third_quintile_retained_customers']}명
#     - 하락: {data['third_quintile_downgraded_customers']}명
#     - 이탈: {data['third_quintile_lost_customers']}명

#     2분위:
#     - 신규: {data['second_quintile_new_customers']}명
#     - 상승: {data['second_quintile_upgraded_customers']}명
#     - 유지: {data['second_quintile_retained_customers']}명
#     - 하락: {data['second_quintile_downgraded_customers']}명
#     - 이탈: {data['second_quintile_lost_customers']}명

#     1분위:
#     - 신규: {data['first_quintile_new_customers']}명
#     - 상승: {data['first_quintile_upgraded_customers']}명
#     - 유지: {data['first_quintile_retained_customers']}명
#     - 하락: {data['first_quintile_downgraded_customers']}명
#     - 이탈: {data['first_quintile_lost_customers']}명

#     ###단계 5: 신규 고객 및 이탈 고객 세부 분석###
#     신규 고객과 이탈 고객에 대한 세부 데이터를 분석하세요:

#     신규 고객:
#     - 총 신규 고객 수: {data['total_new_customers']}명
#     - 5분위로 유입된 신규 고객: {data['fifth_quintile_new_customers']}명
#     - 4분위로 유입된 신규 고객: {data['fourth_quintile_new_customers']}명
#     - 3분위로 유입된 신규 고객: {data['third_quintile_new_customers']}명
#     - 2분위로 유입된 신규 고객: {data['second_quintile_new_customers']}명
#     - 1분위로 유입된 신규 고객: {data['first_quintile_new_customers']}명

#     이탈 고객:
#     - 총 이탈 고객 수: {data['total_lost_customers']}명
#     - 5분위에서 이탈한 고객: {data['fifth_quintile_lost_customers']}명
#     - 4분위에서 이탈한 고객: {data['fourth_quintile_lost_customers']}명
#     - 3분위에서 이탈한 고객: {data['third_quintile_lost_customers']}명
#     - 2분위에서 이탈한 고객: {data['second_quintile_lost_customers']}명
#     - 1분위에서 이탈한 고객: {data['first_quintile_lost_customers']}명

#     신규 고객의 유입 패턴과 이탈 고객의 특성을 분석하여 인사이트를 도출하세요.

#     ###단계 6: 문제점 도출 및 개선 방안 제시###
#     분석한 데이터를 바탕으로 {data['company_name']}의 주요 문제점을 명확히 지적하고, 각 문제에 대한 구체적인 개선 방안을 제시하세요. 다음 사항을 포함하되 이에 국한되지 않습니다:

#     1. 이탈 고객 관리:
#     - 이탈 예측 모델 개발 및 선제적 대응 전략
#     - 1분위 고객 대상 2차 구매 유도 프로그램

#     2. 상위등급 고객 유지:
#     - 등급 하락 방지를 위한 타겟 마케팅 전략
#     - VIP 고객 대상 맞춤형 서비스 강화 방안

#     3. 등급 상승 촉진:
#     - 중위권 고객(3분위, 2분위)의 상위 등급 이동을 위한 인센티브 프로그램
#     - 구매 빈도 및 금액 증대를 위한 크로스셀링/업셀링 전략

#     4. 신규 고객 유치 및 정착:
#     - 신규 고객 유치 전략 및 초기 구매 경험 개선 방안
#     - 신규 고객의 2분위 이상 상승을 위한 프로그램

#     각 전략에 대해 구체적인 실행 방안과 예상되는 효과(고객 충성도 증가, 이탈률 감소, 매출 증대 등)를 제시하세요.

#     ###단계 7: 결론 및 종합 제언###
#     분석 결과를 종합하여 {data['company_name']}의 현재 고객 관리 현황을 간략히 요약하고, 제시한 전략들이 어떻게 전반적인 고객 가치 향상과 매출 증대로 이어질 수 있는지 설명하세요. 또한, 모든 고객 계층에 대한 균형 있는 접근의 중요성을 강조하세요.

#     ###사용자 요청사항###
#     {user_prompt}

#     ###요구사항###
#     - 모든 60개의 데이터셋을 활용하여 종합적이고 깊이 있는 분석을 수행하세요.
#     - 데이터에 기반한 객관적이고 통찰력 있는 분석을 제공하세요.
#     - 실행 가능하고 구체적인 전략을 제시하세요.
#     - 고객 중심의 관점에서 분석과 제안을 작성하세요.
#     - 보고서는 전문적이면서도 이해하기 쉬운 어조로 2000-2500자 내외로 작성하세요.
#     - 각 분위별 특성과 변동 사항을 상세히 비교 분석하여 인사이트를 도출하세요.
#     - 신규 고객 유입과 이탈 고객 현황을 분위별로 상세히 분석하여 전략 수립에 반영하세요.

#     """
#     return base_prompt
def generate_prompt(data, user_prompt):
    base_prompt = f"""
    {{
      "task": "고객 중심의 자본변동표 분석 보고서 작성",
      "company": "{data['company_name']}",
      "year": 2023,
      "total_capital_customers": {data['capital_customers']},
      "data": {{
        "전체_고객_변동": {{
          "신규_유입": {{
            "고객_수": {data['total_new_customers']},
            "비율": {data['total_new_customers'] / data['capital_customers'] * 100:.1f}
          }},
          "이탈": {{
            "고객_수": {data['total_lost_customers']},
            "비율": {data['total_lost_customers'] / data['capital_customers'] * 100:.1f}
          }},
          "등급_상승": {{
            "고객_수": {data['total_upgraded_customers']},
            "비율": {data['total_upgraded_customers'] / data['capital_customers'] * 100:.1f}
          }},
          "등급_하락": {{
            "고객_수": {data['total_downgraded_customers']},
            "비율": {data['total_downgraded_customers'] / data['capital_customers'] * 100:.1f}
          }},
          "등급_유지": {{
            "고객_수": {data['total_retained_customers']},
            "비율": {data['total_retained_customers'] / data['capital_customers'] * 100:.1f}
          }}
        }},
        "고객_변동_비율": {{
          "신규_고객_비율": {data['new_customers_ratio']:.1f},
          "등급_상승률": {data['upgraded_customers_ratio']:.1f},
          "등급_유지율": {data['retained_customers_ratio']:.1f},
          "등급_하락률": {data['downgraded_customers_ratio']:.1f},
          "고객_이탈률": {data['lost_customers_ratio']:.1f},
          "고객_유지율": {data['maintained_customers_ratio']:.1f}
        }},
        "분위별_고객_현황": {{
          "5분위": {{
            "고객_수": {data['fifth_quintile_customers']},
            "비율": {data['fifth_quintile_customers'] / data['capital_customers'] * 100:.1f},
            "매출액": {data['fifth_quintile_sales']:,.0f},
            "매출_비율": {data['fifth_quintile_sales_ratio']:.1f},
            "1인당_평균_매출액": {data['fifth_quintile_avg_sales']:,.0f},
            "변동": {{
              "신규": {data['fifth_quintile_new_customers']},
              "상승": {data['fifth_quintile_upgraded_customers']},
              "유지": {data['fifth_quintile_retained_customers']},
              "하락": {data['fifth_quintile_downgraded_customers']},
              "이탈": {data['fifth_quintile_lost_customers']}
            }}
          }},
          "4분위": {{
            "고객_수": {data['fourth_quintile_customers']},
            "비율": {data['fourth_quintile_customers'] / data['capital_customers'] * 100:.1f},
            "매출액": {data['fourth_quintile_sales']:,.0f},
            "매출_비율": {data['fourth_quintile_sales_ratio']:.1f},
            "1인당_평균_매출액": {data['fourth_quintile_avg_sales']:,.0f},
            "변동": {{
              "신규": {data['fourth_quintile_new_customers']},
              "상승": {data['fourth_quintile_upgraded_customers']},
              "유지": {data['fourth_quintile_retained_customers']},
              "하락": {data['fourth_quintile_downgraded_customers']},
              "이탈": {data['fourth_quintile_lost_customers']}
            }}
          }},
          "3분위": {{
            "고객_수": {data['third_quintile_customers']},
            "비율": {data['third_quintile_customers'] / data['capital_customers'] * 100:.1f},
            "매출액": {data['third_quintile_sales']:,.0f},
            "매출_비율": {data['third_quintile_sales_ratio']:.1f},
            "1인당_평균_매출액": {data['third_quintile_avg_sales']:,.0f},
            "변동": {{
              "신규": {data['third_quintile_new_customers']},
              "상승": {data['third_quintile_upgraded_customers']},
              "유지": {data['third_quintile_retained_customers']},
              "하락": {data['third_quintile_downgraded_customers']},
              "이탈": {data['third_quintile_lost_customers']}
            }}
          }},
          "2분위": {{
            "고객_수": {data['second_quintile_customers']},
            "비율": {data['second_quintile_customers'] / data['capital_customers'] * 100:.1f},
            "매출액": {data['second_quintile_sales']:,.0f},
            "매출_비율": {data['second_quintile_sales_ratio']:.1f},
            "1인당_평균_매출액": {data['second_quintile_avg_sales']:,.0f},
            "변동": {{
              "신규": {data['second_quintile_new_customers']},
              "상승": {data['second_quintile_upgraded_customers']},
              "유지": {data['second_quintile_retained_customers']},
              "하락": {data['second_quintile_downgraded_customers']},
              "이탈": {data['second_quintile_lost_customers']}
            }}
          }},
          "1분위": {{
            "고객_수": {data['first_quintile_customers']},
            "비율": {data['first_quintile_customers'] / data['capital_customers'] * 100:.1f},
            "매출액": {data['first_quintile_sales']:,.0f},
            "매출_비율": {data['first_quintile_sales_ratio']:.1f},
            "1인당_평균_매출액": {data['first_quintile_avg_sales']:,.0f},
            "변동": {{
              "신규": {data['first_quintile_new_customers']},
              "상승": {data['first_quintile_upgraded_customers']},
              "유지": {data['first_quintile_retained_customers']},
              "하락": {data['first_quintile_downgraded_customers']},
              "이탈": {data['first_quintile_lost_customers']}
            }}
          }}
        }}
      }},
      "analysis_steps": [
        {{
          "step": 1,
          "title": "고객변동표 개요",
          "tasks": [
            "고객변동표의 의미와 중요성 설명",
            "고객가치 기준 그룹 분화 방식 소개",
            "5개 고객등급(5분위) 구분 이유 설명",
            "변동 상황 추적 목적 설명",
            "자본고객 기준 설명"
          ]
        }},
        {{
          "step": 2,
          "title": "전체 고객 변동 분석",
          "tasks": [
            "신규 유입, 이탈, 등급 상승/하락/유지 고객 수 분석",
            "고객 변동 비율 분석 (신규, 상승, 유지, 하락, 이탈, 전체 유지)",
            "고객 기반 안정성 평가",
            "성장 가능성 분석"
          ]
        }},
        {{
          "step": 3,
          "title": "분위별 고객 현황 및 매출 분석",
          "tasks": [
            "각 분위별 고객 수 및 비율 분석",
            "분위별 매출 현황 및 기여도 분석",
            "분위별 1인당 평균 매출액 비교",
            "상위 20% 고객의 매출 기여도 특별 분석",
            "하위 고객층 특성 분석"
          ]
        }},
        {{
          "step": 4,
          "title": "분위별 고객 변동 심층 분석",
          "tasks": [
            "각 분위별 신규, 상승, 유지, 하락, 이탈 고객 수 분석",
            "상위등급(5분위, 4분위) 고객의 하락 및 이탈 패턴 분석",
            "1분위 고객의 높은 이탈률 원인 분석",
            "분위 간 고객 이동의 주요 특징 도출"
          ]
        }},
        {{
          "step": 5,
          "title": "신규 고객 및 이탈 고객 세부 분석",
          "tasks": [
            "신규 고객의 분위별 유입 패턴 분석",
            "이탈 고객의 분위별 특성 분석",
            "신규 고객 유입과 이탈 고객 간의 관계 분석",
            "신규 고객 유치 및 이탈 방지를 위한 인사이트 도출"
          ]
        }},
        {{
          "step": 6,
          "title": "문제점 도출 및 개선 방안 제시",
          "tasks": [
            {{
              "area": "이탈 고객 관리",
              "subtasks": [
                "이탈 예측 모델 개발 방안",
                "선제적 대응 전략 수립",
                "1분위 고객 대상 2차 구매 유도 프로그램 설계"
              ]
            }},
            {{
              "area": "상위등급 고객 유지",
              "subtasks": [
                "등급 하락 방지를 위한 타겟 마케팅 전략 수립",
                "VIP 고객 대상 맞춤형 서비스 강화 방안 제시"
              ]
            }},
            {{
              "area": "등급 상승 촉진",
              "subtasks": [
                "중위권 고객 상위 등급 이동 인센티브 프로그램 설계",
                "크로스셀링/업셀링 전략 개발"
              ]
            }},
            {{
              "area": "신규 고객 유치 및 정착",
              "subtasks": [
                "신규 고객 유치 전략 수립",
                "초기 구매 경험 개선 방안 제시",
                "신규 고객의 2분위 이상 상승 프로그램 설계"
              ]
            }}
          ]
        }},
        {{
          "step": 7,
          "title": "결론 및 종합 제언",
          "tasks": [
            "현재 고객 관리 현황 요약",
            "제시된 전략들의 고객 가치 향상 및 매출 증대 효과 설명",
            "모든 고객 계층에 대한 균형 있는 접근의 중요성 강조",
            "장기적 고객 관리 방향성 제시"
          ]
        }}
      ],
      "requirements": [
        "모든 60개 데이터셋을 활용한 종합적이고 깊이 있는 분석 수행",
        "데이터 기반의 객관적이고 통찰력 있는 분석 제공",
        "실행 가능하고 구체적인 전략 제시",
        "고객 중심의 관점에서 분석과 제안 작성",
        "전문적이면서도 이해하기 쉬운 어조로 2000-2500자 내외 작성",
        "각 분위별 특성과 변동 사항의 상세 비교 분석 및 인사이트 도출",
        "신규 고객 유입과 이탈 고객 현황의 분위별 상세 분석 및 전략 반영"
      ],
      "user_request": "{user_prompt}"
    }}
    
    위의 JSON 구조를 기반으로 {data['company_name']}의 2023년 고객변동표를 분석한 보고서를 작성해주세요. 각 분석 단계를 따라가며, 요구사항을 충족시키는 논리적이고 통찰력 있는 보고서를 작성해주시기 바랍니다. 특히 추론 과정을 명확히 보여주며, 각 단계에서 데이터를 기반으로 한 객관적 분석과 그에 따른 주관적 해석을 구분하여 제시해주세요.
    """
    return base_prompt


def generate_customer_equity_change_report(file, user_prompt):
    openai.api_key = os.getenv("OPENAI_API_KEY")
    df = read_excel(file)
    data = process_data(df)
    
    system_message = {
        "role": "system",
        "content": """
        당신은 명확하고 전문적이며 구조화된 비즈니스 보고서를 작성하는 전문가 비즈니스 분석가입니다. 아래 지침을 따라 작성하시오:

        1. 답변은 문단 형식(연속적인 글)으로 작성되어야 하며, 불필요한 경우에는 목록 형식으로 작성하지 마십시오.
        2. 내용을 논리적으로 연결된 문단으로 나누고, 불필요한 목록이나 글머리는 사용하지 않도록 하십시오.
        3. 각 문단은 잘 구성된 생각을 담아야 하며, 섹션 간에 자연스러운 전환이 이루어져야 합니다.
        4. 보고서는 전문적인 어조를 유지해야 하며, 정보를 제공하는 동시에 간결하고 실행 가능한 통찰력을 제공하는 데 중점을 두어야 합니다.
        5. 사용자가 요청한 각 섹션을 다루며, 보고서가 매끄럽게 이어지도록 작성하십시오.
        """
    }

    prompt = generate_prompt(data, user_prompt)
    user_message = {"role": "user", "content": prompt}

    response = openai.ChatCompletion.create(
        model="chatgpt-4o-latest",
        messages=[system_message, user_message],
        max_tokens=2048,
        n=1,
        temperature=0.2,
    )

    return response['choices'][0]['message']['content'].strip()