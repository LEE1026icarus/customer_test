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
        "current_capital_customers": df_ls[df_ls['지표'] == '당기 자본고객 수']['값'].values[0],
        "previous_capital_customers": df_ls[df_ls['지표'] == '전기 자본고객 수']['값'].values[0],
        "large_store_customers": df_ls[df_ls['지표'] == '대형점포 고객 수']['당기'].values[0],
        "large_store_inflow_customers": df_ls[df_ls['지표'] == '대형점포 유입고객 수']['당기'].values[0],
        "large_store_retained_customers": df_ls[df_ls['지표'] == '대형점포 유지고객 수']['당기'].values[0],
        "large_store_outflow_customers": df_ls[df_ls['지표'] == '대형점포 유출고객 수']['당기'].values[0],
        "large_store_flow": df_ls[df_ls['지표'] == '대형점포 고객 흐름']['당기'].values[0],
        "medium_store_customers": df_ls[df_ls['지표'] == '중형점포 고객 수']['당기'].values[0],
        "medium_store_inflow_customers": df_ls[df_ls['지표'] == '중형점포 유입고객 수']['당기'].values[0],
        "medium_store_retained_customers": df_ls[df_ls['지표'] == '중형점포 유지고객 수']['당기'].values[0],
        "medium_store_outflow_customers": df_ls[df_ls['지표'] == '중형점포 유출고객 수']['당기'].values[0],
        "medium_store_flow": df_ls[df_ls['지표'] == '중형점포 고객 흐름']['당기'].values[0],
        "small_store_customers": df_ls[df_ls['지표'] == '소형점포 고객 수']['당기'].values[0],
        "small_store_inflow_customers": df_ls[df_ls['지표'] == '소형점포 유입고객 수']['당기'].values[0],
        "small_store_retained_customers": df_ls[df_ls['지표'] == '소형점포 유지고객 수']['당기'].values[0],
        "small_store_outflow_customers": df_ls[df_ls['지표'] == '소형점포 유출고객 수']['당기'].values[0],
        "small_store_flow": df_ls[df_ls['지표'] == '소형점포 고객 흐름']['당기'].values[0],
        "new_customers": df_ls[df_ls['지표'] == '유입고객 수']['당기'].values[0],
        "retained_customers": df_ls[df_ls['지표'] == '유지고객 수']['당기'].values[0],
        "lost_customers": df_ls[df_ls['지표'] == '유출고객 수']['당기'].values[0],
        "capital_customers_flow": df_ls[df_ls['지표'] == '자본고객 흐름']['당기'].values[0],
    }
    return report_data

# def generate_prompt(data, user_prompt):
#     base_prompt = f"""
#     ###지시사항###

#     당신의 작업은 {data['company_name']}의 2023년 고객흐름표를 분석하여 작성하는 것입니다. 각 점포 규모별 고객 흐름 및 전략적 제언을 포함해주세요.

#     ###단계 1: 고객 흐름 분석###
#     다음 사항을 포함하여 고객 흐름을 분석하세요:
#     - 대형점포 고객 수: {data['large_store_customers']}명
#     - 대형점포 유입 고객 수: {data['large_store_inflow_customers']}명
#     - 대형점포 유지고객 수: {data['large_store_retained_customers']}명
#     - 대형점포 유출 고객 수: {data['large_store_outflow_customers']}명
#     - 대형점포 고객 흐름: {data['large_store_flow']}명
#     - 중형점포 고객 수: {data['medium_store_customers']}명
#     - 중형점포 유입 고객 수: {data['medium_store_inflow_customers']}명
#     - 중형점포 유지고객 수: {data['medium_store_retained_customers']}명
#     - 중형점포 유출 고객 수: {data['medium_store_outflow_customers']}명
#     - 중형점포 고객 흐름: {data['medium_store_flow']}명
#     - 소형점포 고객 수: {data['small_store_customers']}명
#     - 소형점포 유입 고객 수: {data['small_store_inflow_customers']}명
#     - 소형점포 유지고객 수: {data['small_store_retained_customers']}명
#     - 소형점포 유출 고객 수: {data['small_store_outflow_customers']}명
#     - 소형점포 고객 흐름: {data['small_store_flow']}명
#     - 유입고객 수: {data['new_customers']}명
#     - 유지고객 수: {data['retained_customers']}명
#     - 유출고객 수: {data['lost_customers']}명
#     - 자본고객 흐름: {data['capital_customers_flow']}명
#     - 전기 자본고객 수: {data['previous_capital_customers']}명
#     - 당기 자본고객 수: {data['current_capital_customers']}명

#     점포 규모별 자본고객의 흐름을 분석하고, 대형점포에서 자본고객이 감소하고 중/소형점포에서 증가하는 현상의 의미를 설명하세요. 특히 대형점포에서의 고객 유출 원인과 그에 따른 영향에 대해 논의하세요.

#     ###단계 2: 전략적 제언###
#     대형점포의 고객 유출을 방지하고 중/소형점포에서 고객을 유지 및 확대하기 위한 방안을 제시하세요.
#     - 대형점포 고객 유지 방안
#     - 소형점포 고객 확대 전략
#     - 중형점포의 역할 강화 방안

#     각 전략의 실행 방안을 구체적으로 설명하고, 예상되는 효과(예: 고객 유출 감소, 고객 충성도 증가 등)를 서술하세요.

#     ###단계 3: 결론###
#     결론에서는 트라이얼코리아의 현재 고객 흐름의 문제점을 요약하고, 위에서 제시한 전략들이 어떻게 문제 해결에 기여할 수 있는지 설명하세요. 또한 모든 점포 규모에 공평한 접근이 이루어질 수 있도록 주의하세요.

#     ###사용자 요청사항###
#     {user_prompt}

#     ###요구사항###
#     - 자연스럽고 인간적인 어조로 대답하세요.
#     - 객관적인 데이터를 기반으로 결론을 도출하세요.
#     - 보고서는 1000~1300자 내외로 작성하세요.
#     """
#     return base_prompt
def generate_prompt(data, user_prompt):
    base_prompt = f"""
    {{
      "task": "고객흐름표 분석 보고서 작성",
      "company": "{data['company_name']}",
      "year": 2023,
      "data": {{
        "대형점포": {{
          "고객 수": {data['large_store_customers']},
          "유입 고객 수": {data['large_store_inflow_customers']},
          "유지 고객 수": {data['large_store_retained_customers']},
          "유출 고객 수": {data['large_store_outflow_customers']},
          "고객 흐름": {data['large_store_flow']}
        }},
        "중형점포": {{
          "고객 수": {data['medium_store_customers']},
          "유입 고객 수": {data['medium_store_inflow_customers']},
          "유지 고객 수": {data['medium_store_retained_customers']},
          "유출 고객 수": {data['medium_store_outflow_customers']},
          "고객 흐름": {data['medium_store_flow']}
        }},
        "소형점포": {{
          "고객 수": {data['small_store_customers']},
          "유입 고객 수": {data['small_store_inflow_customers']},
          "유지 고객 수": {data['small_store_retained_customers']},
          "유출 고객 수": {data['small_store_outflow_customers']},
          "고객 흐름": {data['small_store_flow']}
        }},
        "전체": {{
          "유입 고객 수": {data['new_customers']},
          "유지 고객 수": {data['retained_customers']},
          "유출 고객 수": {data['lost_customers']},
          "자본고객 흐름": {data['capital_customers_flow']},
          "전기 자본고객 수": {data['previous_capital_customers']},
          "당기 자본고객 수": {data['current_capital_customers']}
        }}
      }},
      "analysis_steps": [
        {{
          "step": 1,
          "title": "고객 흐름 분석",
          "tasks": [
            {{
              "task": "점포 규모별 자본고객 흐름 분석",
              "subtasks": [
                "각 점포 규모별 고객 수, 유입, 유지, 유출 고객 수 비교",
                "자본고객 흐름의 전반적인 추세 파악",
                "점포 규모별 고객 흐름의 차이점 식별"
              ]
            }},
            {{
              "task": "대형점포 고객 감소 현상 분석",
              "subtasks": [
                "대형점포 고객 유출의 주요 원인 추론",
                "중/소형점포 고객 증가와의 연관성 분석",
                "이 현상이 기업에 미치는 잠재적 영향 평가"
              ]
            }},
            {{
              "task": "전체 자본고객 변화 분석",
              "subtasks": [
                "전기 대비 당기 자본고객 수 변화 계산",
                "자본고객 흐름과 전체 고객 흐름의 관계 분석",
                "자본고객 변화가 기업에 미치는 영향 평가"
              ]
            }}
          ]
        }},
        {{
          "step": 2,
          "title": "전략적 제언",
          "tasks": [
            {{
              "task": "대형점포 고객 유지 방안",
              "subtasks": [
                "대형점포 고객 유출 원인에 대응하는 구체적 전략 수립",
                "고객 경험 개선을 위한 아이디어 제시",
                "각 전략의 예상 효과 및 실행 가능성 평가"
              ]
            }},
            {{
              "task": "소형점포 고객 확대 전략",
              "subtasks": [
                "소형점포의 강점 분석 및 활용 방안",
                "소형점포 특화 서비스 또는 상품 제안",
                "소형점포 고객 유치 및 유지 전략 수립"
              ]
            }},
            {{
              "task": "중형점포 역할 강화 방안",
              "subtasks": [
                "중형점포의 현재 위치 및 역할 분석",
                "대형점포와 소형점포 사이에서의 차별화 전략",
                "중형점포 고객 유지 및 확대를 위한 구체적 방안"
              ]
            }}
          ]
        }},
        {{
          "step": 3,
          "title": "결론",
          "tasks": [
            "현재 고객 흐름의 주요 문제점 요약",
            "제안된 전략들이 문제 해결에 기여할 수 있는 방식 설명",
            "모든 점포 규모에 대한 공평한 접근 방식 제시",
            "장기적인 고객 관리 방향성 제안"
          ]
        }}
      ],
      "requirements": [
        "자연스럽고 인간적인 어조로 작성",
        "객관적인 데이터를 기반으로 한 결론 도출",
        "1000~1300자 내외로 보고서 작성",
        "각 단계에서 명확한 추론 과정 제시",
        "논리적 흐름과 일관성 있는 분석 유지",
        "구체적이고 실행 가능한 전략 제안"
      ],
      "user_request": "{user_prompt}"
    }}
    
    위의 JSON 구조를 기반으로 {data['company_name']}의 2023년 고객흐름표를 분석한 보고서를 작성해주세요. 각 분석 단계를 따라가며, 요구사항을 충족시키는 논리적이고 통찰력 있는 보고서를 작성해주시기 바랍니다. 특히 추론 과정을 명확히 보여주며, 각 단계에서 데이터를 기반으로 한 객관적 분석과 그에 따른 주관적 해석을 구분하여 제시해주세요.
    """
    return base_prompt

def generate_customer_cash_flow_report(file, user_prompt):
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