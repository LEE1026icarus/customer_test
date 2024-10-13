import openai
import pandas as pd
import os
from dotenv import load_dotenv
load_dotenv()

# # 데이터로부터 지시사항에 맞게 프롬프트 생성
# def generate_prompt(balance_sheet_result, income_statement_result, cash_flow_result, equity_change_result, user_prompt):
#     base_prompt = f"""
#     ###지시사항###

#     당신의 작업은 특정 회사의 고객 데이터를 분석하여 전략적 제언을 작성하는 것입니다. 아래의 분석 결과와 사용자 요청을 참고하여 보고서를 작성하세요.

#     ###고객 데이터 분석 결과###
#     - 고객상태표(Customer Balance Sheet): 기업이 보유한 총 고객을 자본고객과 부채고객으로 분류하여 고객의 구성 상태를 파악합니다. 분석 결과: {balance_sheet_result}
#     - 고객손익계산서(Customer Income Statement): 자본고객과 부채고객에 대한 손익 계산을 통해 고객 유형별 공헌이익 및 영업이익을 비교합니다. 분석 결과: {income_statement_result}
#     - 고객흐름표(Statement of Customer Flows): 고객을 획득할 수 있는 점포유형을 대형점포, 중형점포, 소형점포로 구분하고 점포규모별 고객의 흐름을 파악합니다. 분석 결과: {cash_flow_result}
#     - 고객변동표(Statement of Changes in Customer): 자본고객을 5분위로 구분하여 각 분위 계층의 현황과 변동사항을 파악합니다. 분석 결과: {equity_change_result}

#     ###전략적 제언###
#     1. 고객 세분화 및 프로그램 강화:
#     현재 멤버십 프로그램의 한계점을 파악하고, 자본고객과 부채고객 간의 차이를 줄이기 위해 멤버십 프로그램을 강화하십시오. 고객 유지를 위해 추가적인 인센티브와 보상 제도를 도입하고, 이탈률이 높은 고객군을 대상으로 고객 락인 전략을 강화하십시오.

#     2. 이탈위험지수 관리 및 대응:
#     고객 이탈률이 높아지는 것을 방지하기 위해 구매행태, 방문 빈도, 구매액 등을 활용한 이탈위험지수를 생성하고, 이를 관리하십시오. 특히, 고가치 고객의 이탈을 방지하는 것이 중요합니다.

#     3. 시계열 데이터 비교를 통한 인사이트 도출:
#     고객 행동을 시계열적으로 분석하여 자본고객 비율의 변화를 파악하고, 이를 통해 전략적 방향성을 설정하십시오. 이러한 분석을 통해 고객 유지 및 확보를 위한 장기적인 전략을 수립하십시오.

#     4. 고객 데이터와 재무 데이터 간 일관성 확보:
#     고객 데이터와 재무 데이터를 비교하여 차이를 최소화하고, 동일한 기준으로 데이터를 관리하십시오. 이를 통해 재무 데이터가 고객 데이터의 보조 역할을 더 효과적으로 수행할 수 있도록 하십시오.

#     5. 기타 전략적 제언:
#     - 비즈니스 애널리틱스를 활용하여 비식별 고객인 부채고객의 수를 더 정확히 추정하고, 고객 전환을 촉진할 수 있는 방안을 마련하십시오.
#     - 매장 방문자 수와 구매 고객 수를 비교하여 고객 전환율을 파악하고, 이를 성과 지표로 활용하여 개선점을 찾아내십시오.

#     ###사용자 요청사항###
#     {user_prompt}

#     ###요구사항###
#     - 자연스럽고 인간적인 방식으로 대답하십시오.
#     - 결과가 편향되지 않고 객관성을 유지하도록 주의하십시오.
#     - 보고서는 줄글 형식으로 작성하되, 내용을 문단으로 구분하여 가독성을 높여주십시오.      
#     - 각 문단의 주제를 명확히 구분하고, 논리적인 흐름을 유지해 주십시오.
#     - 데이터 분석 결과와 전략 제언은 논리적 흐름에 따라 서술해 주십시오.
#     - 보고서 작성 시 논리적 흐름을 유지하며, 단계별로 자연스럽게 전환하십시오.      
#     - 전문적이면서도 이해하기 쉬운 언어를 사용하여 작성해 주십시오.   
#     """
    
#     return base_prompt
def generate_prompt(balance_sheet_result, income_statement_result, cash_flow_result, equity_change_result, user_prompt):
    base_prompt = f"""
    {{
      "task": "고객 데이터 분석 및 전략적 제언 보고서 작성",
      "data": {{
        "고객상태표": "{balance_sheet_result}",
        "고객손익계산서": "{income_statement_result}",
        "고객흐름표": "{cash_flow_result}",
        "고객변동표": "{equity_change_result}"
      }},
      "analysis_steps": [
        {{
          "step": 1,
          "title": "고객 데이터 종합 분석",
          "tasks": [
            {{
              "task": "고객상태표 분석",
              "subtasks": [
                "자본고객과 부채고객의 비율 파악",
                "고객 구성의 변화 추세 분석",
                "자본고객과 부채고객의 특성 비교"
              ]
            }},
            {{
              "task": "고객손익계산서 분석",
              "subtasks": [
                "자본고객과 부채고객의 공헌이익 비교",
                "고객 유형별 영업이익 분석",
                "고객 유형별 수익성 차이의 원인 파악"
              ]
            }},
            {{
              "task": "고객흐름표 분석",
              "subtasks": [
                "점포 규모별 고객 획득 현황 파악",
                "고객 이동 패턴 분석",
                "점포 규모와 고객 유지율의 관계 분석"
              ]
            }},
            {{
              "task": "고객변동표 분석",
              "subtasks": [
                "자본고객의 5분위 구분에 따른 현황 파악",
                "각 분위별 고객 변동 추이 분석",
                "상위 분위와 하위 분위 간의 차이점 도출"
              ]
            }}
          ]
        }},
        {{
          "step": 2,
          "title": "전략적 제언 도출",
          "tasks": [
            {{
              "area": "고객 세분화 및 프로그램 강화",
              "subtasks": [
                "현재 멤버십 프로그램의 한계점 파악",
                "자본고객과 부채고객 간 차이를 줄이기 위한 방안 제시",
                "고객 유지를 위한 인센티브 및 보상 제도 설계",
                "이탈률이 높은 고객군 대상 락인 전략 수립"
              ]
            }},
            {{
              "area": "이탈위험지수 관리 및 대응",
              "subtasks": [
                "이탈위험지수 생성 방법론 제안",
                "고가치 고객의 이탈 방지를 위한 특별 관리 방안 수립",
                "이탈위험지수에 따른 차별화된 고객 관리 전략 제시"
              ]
            }},
            {{
              "area": "시계열 데이터 비교를 통한 인사이트 도출",
              "subtasks": [
                "고객 행동의 시계열적 분석 방법 제안",
                "자본고객 비율 변화에 따른 전략적 방향성 설정",
                "장기적 고객 유지 및 확보 전략 수립"
              ]
            }},
            {{
              "area": "고객 데이터와 재무 데이터 간 일관성 확보",
              "subtasks": [
                "고객 데이터와 재무 데이터 간 차이점 분석",
                "데이터 관리 기준 통일 방안 제시",
                "통합 데이터 관리 시스템 구축 제안"
              ]
            }},
            {{
              "area": "추가 전략적 제언",
              "subtasks": [
                "비즈니스 애널리틱스를 활용한 부채고객 추정 방법 제안",
                "고객 전환율 향상을 위한 전략 수립",
                "매장 방문자 수와 구매 고객 수 비교 분석 및 개선점 도출"
              ]
            }}
          ]
        }},
        {{
          "step": 3,
          "title": "결론 및 종합 제언",
          "tasks": [
            "주요 분석 결과 요약",
            "핵심 전략적 제언 정리",
            "제안된 전략들의 예상 효과 및 우선순위 제시",
            "장기적 고객 관리 방향성 제안"
          ]
        }}
      ],
      "requirements": [
        "자연스럽고 인간적인 어조로 보고서 작성",
        "객관적이고 편향되지 않은 분석 결과 제시",
        "논리적 흐름을 유지하며 단계별로 자연스럽게 전환",
        "전문적이면서도 이해하기 쉬운 언어 사용",
        "각 분석 단계에서 명확한 추론 과정 제시",
        "데이터 기반의 객관적 분석과 주관적 해석의 명확한 구분",
        "실행 가능하고 구체적인 전략 제안"
      ],
      "user_request": "{user_prompt}"
    }}
    
    위의 JSON 구조를 기반으로 고객 데이터 분석 및 전략적 제언 보고서를 작성해주세요. 각 분석 단계를 따라가며, 요구사항을 충족시키는 논리적이고 통찰력 있는 보고서를 작성해주시기 바랍니다. 특히 추론 과정을 명확히 보여주며, 각 단계에서 데이터를 기반으로 한 객관적 분석과 그에 따른 주관적 해석을 구분하여 제시해주세요. 보고서는 줄글 형식으로 작성하되, 내용을 문단으로 구분하여 가독성을 높여주시기 바랍니다.
    """
    return base_prompt

# 전략적 제언 보고서 생성 함수
def generate_strategic_recommendations_report(balance_sheet_result, income_statement_result, cash_flow_result, equity_change_result,user_prompt):
    try:
        # 데이터를 처리하여 지시사항에 맞는 프롬프트 생성
        prompt = generate_prompt(balance_sheet_result, income_statement_result, cash_flow_result, equity_change_result,user_prompt)

        # OpenAI API 호출을 위한 시스템 메시지 설정
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

        # 사용자 메시지 설정
        user_message = {"role": "user", "content": prompt}

        # OpenAI API 호출
        openai.api_key = os.getenv("OPENAI_API_KEY")

        response = openai.ChatCompletion.create(
            model="chatgpt-4o-latest",
            messages=[system_message, user_message],
            max_tokens=2048,
            n=1,
            temperature=0.2,
        )

        return response['choices'][0]['message']['content'].strip()
    
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        return f"전략적 제언 생성 중 오류가 발생했습니다: {str(e)}\n\n상세 오류:\n{error_details}"
