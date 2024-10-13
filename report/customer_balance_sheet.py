import openai
import pandas as pd
import io
import os
from dotenv import load_dotenv
load_dotenv()



def read_excel(file):
    # UploadedFile 객체를 바이트 스트림으로 변환
    bytes_data = file.getvalue()
    
    # 바이트 스트림을 io.BytesIO 객체로 변환
    excel_data = io.BytesIO(bytes_data)
    
    # pandas로 엑셀 파일 읽기
    df_ls = pd.read_excel(excel_data, sheet_name='LS')
    return df_ls

def process_data(df_ls):
    report_data = {
        "company_name": df_ls[df_ls['지표'] == '기업명']['당기'].values[0],
        "total_customers": df_ls[df_ls['지표'] == '자산고객 수']['당기'].values[0],
        "capital_customers": df_ls[df_ls['지표'] == '당기 자본고객 수']['당기'].values[0],
        "capital_customer_ratio": df_ls[df_ls['지표'] == '자본고객 비율']['당기'].values[0],
        "debt_customers": df_ls[df_ls['지표'] == '부채고객 수']['당기'].values[0],
        "debt_customer_ratio": df_ls[df_ls['지표'] == '부채고객 비율']['당기'].values[0],
        "total_sales": df_ls[df_ls['지표'] == '자산고객 총 매출액']['당기'].values[0],
        "capital_sales": df_ls[df_ls['지표'] == '자본고객 총 매출액']['당기'].values[0],
        "capital_sales_ratio": df_ls[df_ls['지표'] == '자본고객 매출비율']['당기'].values[0],
        "debt_sales": df_ls[df_ls['지표'] == '부채고객 총 매출액']['당기'].values[0],
        "debt_sales_ratio": df_ls[df_ls['지표'] == '부채고객 매출비율']['당기'].values[0],
        "capital_avg_sales": df_ls[df_ls['지표'] == '자본고객 평균 매출액']['당기'].values[0],
        "debt_avg_sales": df_ls[df_ls['지표'] == '부채고객 평균 매출액']['당기'].values[0],
        "asset_avg_sales": df_ls[df_ls['지표'] == '자산고객 평균 매출액']['당기'].values[0],
        "current_capital_customers": df_ls[df_ls['지표'] == '당기 자본고객 수']['당기'].values[0],
    }
    return report_data

# def generate_prompt(data, user_prompt):
#     base_prompt = f"""
#     ###지시사항###

#     당신의 작업은 {data['company_name']}의 2023년 고객상태표 관점의 소견을 제시하는 분석입니다. 각 단계에서 자세한 추론 과정을 보여주며, 논리적 흐름을 유지하세요.

#     ###단계 1: 고객 데이터 분석###
#     다음 데이터를 분석하세요:
#     - 총 고객 수: {data['total_customers']}
#     - 자본고객 수: {data['capital_customers']} ({data['capital_customer_ratio']*100:.1f}%)
#     - 부채고객 수: {data['debt_customers']} ({data['debt_customer_ratio']*100:.1f}%)
#     - 전체 매출액: {float(data['total_sales'].replace('원', '').replace(',', '')) / 1e8:.1f}억원
#     - 자본고객 매출액: {float(data['capital_sales'].replace('원', '').replace(',', '')) / 1e8:.1f}억원 ({data['capital_sales_ratio']*100:.1f}%)
#     - 부채고객 매출액: {float(data['debt_sales'].replace('원', '').replace(',', '')) / 1e8:.1f}억원 ({data['debt_sales_ratio']*100:.1f}%)
#     - 자본고객 평균 매출액: {float(data['capital_avg_sales'].replace('원', '').replace(',', '')) / 1e4:.1f}만원
#     - 부채고객 평균 매출액: {float(data['debt_avg_sales'].replace('원', '').replace(',', '')) / 1e4:.1f}만원
#     - 자산고객 평균 매출액: {float(data['asset_avg_sales'].replace('원', '').replace(',', '')) / 1e4:.1f}만원
#     - 당기 자본고객 수: {data['current_capital_customers']}
#     - 자본고객 매출 비율: {data['capital_sales_ratio']*100:.1f}%
#     - 부채고객 매출 비율: {data['debt_sales_ratio']*100:.1f}%
#     - 자본고객 비율: {data['capital_customer_ratio']*100:.1f}%

#     추론 과정:
#     1. 자본고객과 부채고객의 비율을 비교하세요.
#     2. 각 고객 그룹의 매출 기여도를 분석하세요.
#     3. 평균 매출액의 차이가 의미하는 바를 고찰하세요.
#     4. 이 데이터가 기업의 전반적인 재무 건전성에 어떤 영향을 미치는지 추론하세요.

#     ###단계 2: 자본고객 유지 및 확대 전략###
#     추론 과정:
#     1. 자본고객의 중요성을 설명하세요.
#     2. 맞춤형 서비스 제공 전략을 구체화하세요:
#        a) 어떤 종류의 맞춤형 서비스가 효과적일지 고려하세요.
#        b) 이 서비스가 고객 충성도에 미칠 영향을 예측하세요.
#     3. 추천 프로그램 도입 전략을 구체화하세요:
#        a) 추천 프로그램의 구조를 설계하세요.
#        b) 이 프로그램이 신규 고객 유입에 미칠 영향을 예측하세요.
#     4. 각 전략의 예상 효과를 수치화하여 제시하세요.

#     ###단계 3: 부채고객의 자본고객 전환 전략###
#     추론 과정:
#     1. 부채고객의 특성을 분석하세요.
#     2. 부채고객이 자본고객이 되지 못하는 주요 원인을 추론하세요.
#     3. 이러한 원인을 해결할 수 있는 전략을 고안하세요.
#     4. 각 전략의 실행 방안과 예상 효과를 구체적으로 설명하세요.

#     ###단계 4: 부채고객 매출 증대 전략###
#     추론 과정:
#     1. 현재 부채고객의 매출 특성을 분석하세요.
#     2. 부채고객의 매출을 증가시킬 수 있는 요인들을 나열하세요.
#     3. 각 요인에 대응하는 구체적인 전략을 수립하세요.
#     4. 이 전략들이 자본고객으로의 전환에 미칠 수 있는 긍정적 영향을 분석하세요.

#     ###단계 5: 편향되지 않은 결론 도출###
#     추론 과정:
#     1. 각 단계에서 도출된 주요 결과를 정리하세요.
#     2. 이 결과들이 특정 고객 그룹에 편향되어 있는지 검토하세요.
#     3. 편향이 발견된다면, 이를 조정할 수 있는 방안을 제시하세요.
#     4. 모든 고객 그룹을 공평하게 고려한 최종 결론을 도출하세요.

#     ###사용자 요청사항###
#     {user_prompt}

#     ###요구사항###
#     - 각 단계에서 추론 과정을 명확히 보여주세요.
#     - 결론에 이르기까지의 논리적 연결고리를 강화하세요.
#     - 데이터를 기반으로 한 객관적인 분석과 주관적 해석을 구분하여 제시하세요.
#     - 각 전략의 예상 효과를 구체적인 수치나 비율로 제시하려 노력하세요.
#     - 전체 보고서는 1000~1300자 내외로 작성하되, 각 단계별 추론 과정이 명확히 드러나도록 하세요.
#     - 최종 결론에서는 각 단계의 추론 결과를 종합하여 일관성 있는 전략을 제시하세요.
#     """
#     return base_prompt
def generate_prompt(data, user_prompt):
    base_prompt = f"""
    {{
      "task": "분석 보고서 작성",
      "company": "{data['company_name']}",
      "year": 2023,
      "data": {{
        "총 고객 수": {data['total_customers']},
        "자본고객 수": {data['capital_customers']},
        "자본고객 비율": {data['capital_customer_ratio']*100:.1f},
        "부채고객 수": {data['debt_customers']},
        "부채고객 비율": {data['debt_customer_ratio']*100:.1f},
        "전체 매출액": {float(data['total_sales'].replace('원', '').replace(',', '')) / 1e8:.1f},
        "자본고객 매출액": {float(data['capital_sales'].replace('원', '').replace(',', '')) / 1e8:.1f},
        "자본고객 매출 비율": {data['capital_sales_ratio']*100:.1f},
        "부채고객 매출액": {float(data['debt_sales'].replace('원', '').replace(',', '')) / 1e8:.1f},
        "부채고객 매출 비율": {data['debt_sales_ratio']*100:.1f},
        "자본고객 평균 매출액": {float(data['capital_avg_sales'].replace('원', '').replace(',', '')) / 1e4:.1f},
        "부채고객 평균 매출액": {float(data['debt_avg_sales'].replace('원', '').replace(',', '')) / 1e4:.1f},
        "자산고객 평균 매출액": {float(data['asset_avg_sales'].replace('원', '').replace(',', '')) / 1e4:.1f},
        "당기 자본고객 수": {data['current_capital_customers']}
      }},
      "analysis_steps": [
        {{
          "step": 1,
          "title": "고객 데이터 분석",
          "tasks": [
            "자본고객과 부채고객의 비율 비교",
            "각 고객 그룹의 매출 기여도 분석",
            "평균 매출액 차이의 의미 고찰",
            "기업의 재무 건전성에 미치는 영향 추론"
          ]
        }},
        {{
          "step": 2,
          "title": "자본고객 유지 및 확대 전략",
          "tasks": [
            "자본고객의 중요성 설명",
            {{
              "strategy": "맞춤형 서비스 제공",
              "subtasks": [
                "효과적인 맞춤형 서비스 유형 고려",
                "고객 충성도에 미칠 영향 예측"
              ]
            }},
            {{
              "strategy": "추천 프로그램 도입",
              "subtasks": [
                "추천 프로그램 구조 설계",
                "신규 고객 유입에 미칠 영향 예측"
              ]
            }},
            "각 전략의 예상 효과 수치화"
          ]
        }},
        {{
          "step": 3,
          "title": "부채고객의 자본고객 전환 전략",
          "tasks": [
            "부채고객의 특성 분석",
            "자본고객 전환 장애 요인 추론",
            "장애 요인 해결 전략 고안",
            "전략별 실행 방안과 예상 효과 설명"
          ]
        }},
        {{
          "step": 4,
          "title": "부채고객 매출 증대 전략",
          "tasks": [
            "부채고객의 현재 매출 특성 분석",
            "매출 증가 요인 나열",
            "요인별 구체적 전략 수립",
            "자본고객 전환에 미칠 긍정적 영향 분석"
          ]
        }},
        {{
          "step": 5,
          "title": "편향되지 않은 결론 도출",
          "tasks": [
            "각 단계의 주요 결과 정리",
            "결과의 편향성 검토",
            "편향 조정 방안 제시",
            "모든 고객 그룹을 공평하게 고려한 최종 결론 도출"
          ]
        }}
      ],
      "requirements": [
        "각 단계에서 추론 과정을 명확히 보여줄 것",
        "결론까지의 논리적 연결고리 강화",
        "객관적 분석과 주관적 해석 구분",
        "전략의 예상 효과를 구체적 수치나 비율로 제시",
        "전체 보고서는 1000~1300자 내외로 작성",
        "각 단계별 추론 과정이 명확히 드러나도록 할 것",
        "최종 결론에서 각 단계의 추론 결과를 종합하여 일관성 있는 전략 제시"
      ],
      "user_request": "{user_prompt}"
    }}
    
    위의 JSON 구조를 기반으로 {data['company_name']}의 2023년 고객상태표 관점의 소견을 제시하는 분석 보고서를 작성해주세요. 각 분석 단계를 따라가며, 요구사항을 충족시키는 논리적이고 통찰력 있는 보고서를 작성해주시기 바랍니다.
    """
    return base_prompt

def generate_customer_balance_sheet_report(file, user_prompt):
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