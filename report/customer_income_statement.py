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
        "자산고객 총 매출액": df_ls[df_ls['지표'] == '자산고객 총 매출액']['당기'].values[0],
        "총 비용": df_ls[df_ls['지표'] == '총 비용']['당기'].values[0],
        "변동비": df_ls[df_ls['지표'] == '변동비']['당기'].values[0],
        "변동비율": float(df_ls[df_ls['지표'] == '변동비율']['당기'].values[0]),
        "공헌이익": df_ls[df_ls['지표'] == '공헌이익']['당기'].values[0],
        "공헌이익률": float(df_ls[df_ls['지표'] == '공헌이익률']['당기'].values[0]),
        "고정비": df_ls[df_ls['지표'] == '고정비']['당기'].values[0],
        "고정비율": float(df_ls[df_ls['지표'] == '고정비율']['당기'].values[0]),
        "영업이익": df_ls[df_ls['지표'] == '영업이익']['당기'].values[0],
        "영업이익률": float(df_ls[df_ls['지표'] == '영업이익률']['당기'].values[0]),
        "1인 평균 공헌이익": df_ls[df_ls['지표'] == '1인 평균 공헌이익']['당기'].values[0],
        "1인 평균 영업이익": df_ls[df_ls['지표'] == '1인 평균 영업이익']['당기'].values[0],
        "자본고객 총 매출액": df_ls[df_ls['지표'] == '자본고객 총 매출액']['당기'].values[0],
        "자본고객 변동비": df_ls[df_ls['지표'] == '자본고객 변동비']['당기'].values[0],
        "자본고객 변동비율": float(df_ls[df_ls['지표'] == '자본고객 변동비율']['당기'].values[0]),
        "자본고객 공헌이익": df_ls[df_ls['지표'] == '자본고객 공헌이익']['당기'].values[0],
        "자본고객 공헌이익률": float(df_ls[df_ls['지표'] == '자본고객 공헌이익률']['당기'].values[0]),
        "자본고객 고정비": df_ls[df_ls['지표'] == '자본고객 고정비']['당기'].values[0],
        "자본고객 고정비율": float(df_ls[df_ls['지표'] == '자본고객 고정비율']['당기'].values[0]),
        "자본고객 영업이익": df_ls[df_ls['지표'] == '자본고객 영업이익']['당기'].values[0],
        "자본고객 영업이익률": float(df_ls[df_ls['지표'] == '자본고객 영업이익률']['당기'].values[0]),
        "자본고객 1인 평균 공헌이익": df_ls[df_ls['지표'] == '자본고객 1인 평균 공헌이익']['당기'].values[0],
        "자본고객 1인 평균 영업이익": df_ls[df_ls['지표'] == '자본고객 1인 평균 영업이익']['당기'].values[0],
        "부채고객 총 매출액": df_ls[df_ls['지표'] == '부채고객 총 매출액']['당기'].values[0],
        "부채고객 변동비": df_ls[df_ls['지표'] == '부채고객 변동비']['당기'].values[0],
        "부채고객 변동비율": float(df_ls[df_ls['지표'] == '부채고객 변동비율']['당기'].values[0]),
        "부채고객 공헌이익": df_ls[df_ls['지표'] == '부채고객 공헌이익']['당기'].values[0],
        "부채고객 공헌이익률": float(df_ls[df_ls['지표'] == '부채고객 공헌이익률']['당기'].values[0]),
        "부채고객 고정비": df_ls[df_ls['지표'] == '부채고객 고정비']['당기'].values[0],
        "부채고객 고정비율": float(df_ls[df_ls['지표'] == '부채고객 고정비율']['당기'].values[0]),
        "부채고객 영업이익": df_ls[df_ls['지표'] == '부채고객 영업이익']['당기'].values[0],
        "부채고객 영업이익률": float(df_ls[df_ls['지표'] == '부채고객 영업이익률']['당기'].values[0]),
        "부채고객 1인 평균 공헌이익": df_ls[df_ls['지표'] == '부채고객 1인 평균 공헌이익']['당기'].values[0],
        "부채고객 1인 평균 영업이익": df_ls[df_ls['지표'] == '부채고객 1인 평균 영업이익']['당기'].values[0]
    }
    return report_data

def generate_prompt(data, user_prompt):
    base_prompt = f"""
    {{
      "task": "고객손익계산서 분석 보고서 작성",
      "company": "{data['company_name']}",
      "year": 2023,
      "data": {{
        "전체_고객": {{
          "자산고객_총_매출액": {int(data['자산고객 총 매출액'].replace('원', '').replace(',', '')) / 1e8:.1f},
          "총_비용": {int(data['총 비용'].replace('원', '').replace(',', '')) / 1e8:.1f},
          "변동비": {{
            "금액": {int(data['변동비'].replace('원', '').replace(',', '')) / 1e8:.1f},
            "비율": {data['변동비율'] * 100:.1f}
          }},
          "공헌이익": {{
            "금액": {int(data['공헌이익'].replace('원', '').replace(',', '')) / 1e8:.1f},
            "비율": {data['공헌이익률'] * 100:.1f}
          }},
          "고정비": {{
            "금액": {int(data['고정비'].replace('원', '').replace(',', '')) / 1e8:.1f},
            "비율": {data['고정비율'] * 100:.1f}
          }},
          "영업이익": {{
            "금액": {int(data['영업이익'].replace('원', '').replace(',', '')) / 1e8:.1f},
            "비율": {data['영업이익률'] * 100:.1f}
          }},
          "1인_평균_공헌이익": {int(data['1인 평균 공헌이익'].replace('원', '').replace(',', ''))},
          "1인_평균_영업이익": {int(data['1인 평균 영업이익'].replace('원', '').replace(',', ''))}
        }},
        "자본고객": {{
          "총_매출액": {int(data['자본고객 총 매출액'].replace('원', '').replace(',', '')) / 1e8:.1f},
          "변동비": {{
            "금액": {int(data['자본고객 변동비'].replace('원', '').replace(',', '')) / 1e8:.1f},
            "비율": {data['자본고객 변동비율'] * 100:.1f}
          }},
          "공헌이익": {{
            "금액": {int(data['자본고객 공헌이익'].replace('원', '').replace(',', '')) / 1e8:.1f},
            "비율": {data['자본고객 공헌이익률'] * 100:.1f}
          }},
          "고정비": {{
            "금액": {int(data['자본고객 고정비'].replace('원', '').replace(',', '')) / 1e8:.1f},
            "비율": {data['자본고객 고정비율'] * 100:.1f}
          }},
          "영업이익": {{
            "금액": {int(data['자본고객 영업이익'].replace('원', '').replace(',', '')) / 1e8:.1f},
            "비율": {data['자본고객 영업이익률'] * 100:.1f}
          }},
          "1인_평균_공헌이익": {int(data['자본고객 1인 평균 공헌이익'].replace('원', '').replace(',', ''))},
          "1인_평균_영업이익": {int(data['자본고객 1인 평균 영업이익'].replace('원', '').replace(',', ''))}
        }},
        "부채고객": {{
          "총_매출액": {int(data['부채고객 총 매출액'].replace('원', '').replace(',', '')) / 1e8:.1f},
          "변동비": {{
            "금액": {int(data['부채고객 변동비'].replace('원', '').replace(',', '')) / 1e8:.1f},
            "비율": {data['부채고객 변동비율'] * 100:.1f}
          }},
          "공헌이익": {{
            "금액": {int(data['부채고객 공헌이익'].replace('원', '').replace(',', '')) / 1e8:.1f},
            "비율": {data['부채고객 공헌이익률'] * 100:.1f}
          }},
          "고정비": {{
            "금액": {int(data['부채고객 고정비'].replace('원', '').replace(',', '')) / 1e8:.1f},
            "비율": {data['부채고객 고정비율'] * 100:.1f}
          }},
          "영업이익": {{
            "금액": {int(data['부채고객 영업이익'].replace('원', '').replace(',', '')) / 1e8:.1f},
            "비율": {data['부채고객 영업이익률'] * 100:.1f}
          }},
          "1인_평균_공헌이익": {int(data['부채고객 1인 평균 공헌이익'].replace('원', '').replace(',', ''))},
          "1인_평균_영업이익": {int(data['부채고객 1인 평균 영업이익'].replace('원', '').replace(',', ''))}
        }}
      }},
      "analysis_steps": [
        {{
          "step": 1,
          "title": "전체 고객 손익 분석",
          "tasks": [
            "전체 매출액과 비용 구조 분석",
            "변동비와 고정비의 비율 평가",
            "공헌이익과 영업이익의 규모 및 비율 분석",
            "1인당 평균 공헌이익과 영업이익 평가"
          ]
        }},
        {{
          "step": 2,
          "title": "자본고객 그룹 분석",
          "tasks": [
            "자본고객의 매출 기여도 분석",
            "자본고객의 비용 구조 평가",
            "자본고객의 공헌이익과 영업이익 분석",
            "자본고객의 1인당 평균 이익 평가",
            "전체 고객 대비 자본고객의 수익성 비교"
          ]
        }},
        {{
          "step": 3,
          "title": "부채고객 그룹 분석",
          "tasks": [
            "부채고객의 매출 기여도 분석",
            "부채고객의 비용 구조 평가",
            "부채고객의 공헌이익과 영업이익 분석",
            "부채고객의 1인당 평균 이익 평가",
            "전체 고객 및 자본고객 대비 부채고객의 수익성 비교"
          ]
        }},
        {{
          "step": 4,
          "title": "전략적 제언",
          "tasks": [
            {{
              "area": "부채고객 비율 감소",
              "subtasks": [
                "부채고객이 수익성에 미치는 영향 분석",
                "자본고객 중심의 마케팅 전략 수립",
                "부채고객의 자본고객 전환 방안 제시"
              ]
            }},
            {{
              "area": "고정비 항목 점검 및 최적화",
              "subtasks": [
                "고정비 항목 세부 분석",
                "불필요한 지출 식별 및 감축 방안",
                "일부 고정비의 변동비 전환 가능성 검토"
              ]
            }},
            {{
              "area": "포인트 적립 및 사용 회계처리 재검토",
              "subtasks": [
                "현행 포인트 회계처리 방식의 문제점 분석",
                "대안적 회계처리 방식 제안",
                "새로운 방식 도입 시 예상되는 효과 분석"
              ]
            }}
          ]
        }},
        {{
          "step": 5,
          "title": "결론",
          "tasks": [
            "현재 고객 포트폴리오의 주요 문제점 요약",
            "제안된 전략들의 예상 효과 종합",
            "모든 고객 유형에 대한 균형 있는 접근 방식 제시",
            "장기적 고객 수익성 개선을 위한 로드맵 제안"
          ]
        }}
      ],
      "requirements": [
        "자연스럽고 인간적인 어조로 보고서 작성",
        "객관적인 데이터를 기반으로 한 결론 도출",
        "1000~1300자 내외의 보고서 작성",
        "각 분석 단계에서 명확한 추론 과정 제시",
        "데이터 기반의 객관적 분석과 주관적 해석의 명확한 구분",
        "실행 가능하고 구체적인 전략 제안"
      ],
      "user_request": "{user_prompt}"
    }}
    
    위의 JSON 구조를 기반으로 {data['company_name']}의 2023년 고객손익계산서를 분석한 보고서를 작성해주세요. 각 분석 단계를 따라가며, 요구사항을 충족시키는 논리적이고 통찰력 있는 보고서를 작성해주시기 바랍니다. 특히 추론 과정을 명확히 보여주며, 각 단계에서 데이터를 기반으로 한 객관적 분석과 그에 따른 주관적 해석을 구분하여 제시해주세요.
    """
    return base_prompt

def generate_customer_income_statement_report(file, user_prompt):
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