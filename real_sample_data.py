# 실제 G임팩트 분석 데이터 기반 샘플

REAL_SAMPLE_DATA = {
    "step_2_1_pestel": {
        "analysis_meta": {
            "company": "G임팩트",
            "analysis_date": "2026-01-05"
        },
        "executive_summary": "2026년 G임팩트는 정부의 AI·딥테크 육성 정책과 지역 소멸 대응 기조에 힘입어 강력한 성장 기회를 맞이했습니다. SLM 기술을 통한 내부 리소스 문제 해결이 가능하나, 벤처 투자 침체와 AI 규제 강화, 그리고 심각한 내부 번아웃이 주요 위협입니다.",
        "pestel": {
            "political": {
                "summary": "정부의 AI 육성 및 지역 균형 발전 정책은 호재이나, 높은 B2G 의존도는 리스크.",
                "issues": [
                    {"id": "P1", "name": "AI·딥테크 스타트업 육성 정책", "description": "2030년까지 AI 스타트업 집중 육성 및 자금 지원", "impact_score": "5", "urgency_score": "4", "classification": "기회"},
                    {"id": "P2", "name": "비수도권 정책금융 확대", "description": "지방공급 확대 목표제로 비수도권 자금 배정 증가", "impact_score": "4", "urgency_score": "3", "classification": "기회"},
                    {"id": "P3", "name": "높은 B2G 매출 의존도", "description": "80%에 달하는 정부지원금 의존율에 따른 정책 변동 리스크", "impact_score": "5", "urgency_score": "5", "classification": "위협"}
                ],
                "opportunities": [{"rank": 1, "factor": "지역 주도형 AI 사업 확대", "action": "광주·전남 지자체 특화 AI 바우처 사업 제안서 선제적 제출"}],
                "threats": [{"rank": 1, "factor": "정권/지자체 정책 기조 변화", "mitigation": "B2B 매출 비중 확대 (2026년 목표 30%)"}]
            },
            "economic": {
                "summary": "민간 투자 위축을 공공 모태펀드와 성장하는 SaaS 시장 수요로 방어해야 합니다.",
                "issues": [
                    {"id": "E1", "name": "AI SaaS 시장 고성장", "description": "글로벌 AI SaaS 시장 연 38% 성장", "impact_score": "4", "urgency_score": "3", "classification": "기회"},
                    {"id": "E2", "name": "벤처투자 시장 침체", "description": "초기 스타트업 투자 건수 및 규모 급감", "impact_score": "5", "urgency_score": "4", "classification": "위협"},
                    {"id": "E3", "name": "정부 모태펀드 확대", "description": "1.9조 원 규모 모태펀드 출자", "impact_score": "4", "urgency_score": "3", "classification": "기회"}
                ],
                "opportunities": [{"rank": 1, "factor": "SaaS 바우처 사업", "action": "K-비대면 바우처 등 공급기업 자격 유지"}],
                "threats": [{"rank": 1, "factor": "민간 시장의 지불 의사 하락", "mitigation": "성과 연동형 과금 모델 검토"}]
            },
            "social": {
                "summary": "지역 소멸 위기는 사업의 명분을 주지만, 내부 구성원의 번아웃은 운영의 위기입니다.",
                "issues": [
                    {"id": "S1", "name": "지방 소멸 위험 심화", "description": "전남 지역 소멸위험지수 0.32", "impact_score": "5", "urgency_score": "3", "classification": "기회"},
                    {"id": "S2", "name": "MZ세대의 가치소비", "description": "환경, 윤리 등을 중시하는 소비 트렌드", "impact_score": "3", "urgency_score": "2", "classification": "기회"},
                    {"id": "S3", "name": "내부 구성원 번아웃", "description": "CEO 휴일근무 6일, 평균 연장근로 12.5시간", "impact_score": "5", "urgency_score": "5", "classification": "위협"}
                ],
                "opportunities": [{"rank": 1, "factor": "지역 재생 프로그램 수요", "action": "관계인구 유입 및 로컬 크리에이터 육성 사업 확대"}],
                "threats": [{"rank": 1, "factor": "핵심 인재 이탈 및 번아웃", "mitigation": "자사 웰니스 솔루션 도입 및 강제 휴식 제도"}]
            },
            "technological": {
                "summary": "SLM 기술 도입은 개발 인력 부족과 비용 문제를 동시에 해결할 수 있습니다.",
                "issues": [
                    {"id": "T1", "name": "SLM(소형언어모델) 확산", "description": "저비용, 고효율, 온프레미스 구축 가능", "impact_score": "5", "urgency_score": "4", "classification": "기회"},
                    {"id": "T2", "name": "기업 생성형 AI 도입 가속", "description": "국내 기업 AI 도입률 85% 전망", "impact_score": "4", "urgency_score": "3", "classification": "기회"},
                    {"id": "T3", "name": "AI 신뢰성 및 보안 이슈", "description": "환각 현상 및 데이터 보안 우려 증가", "impact_score": "3", "urgency_score": "3", "classification": "위협"}
                ],
                "opportunities": [{"rank": 1, "factor": "기술 효율화 (SLM)", "action": "오픈소스 SLM 모델 파인튜닝 착수"}],
                "threats": [{"rank": 1, "factor": "기술 격차 심화", "mitigation": "핵심 기능에 집중하여 기술 깊이 확보"}]
            },
            "environmental": {
                "summary": "공급망 ESG 실사는 B2B SaaS 확장의 강력한 트리거입니다.",
                "issues": [
                    {"id": "En1", "name": "공급망 ESG 실사 확산", "description": "대기업 협력사에 대한 ESG 데이터 제출 요구 증가", "impact_score": "4", "urgency_score": "4", "classification": "기회"},
                    {"id": "En2", "name": "ESG 공시 의무화 지연", "description": "2026년 이후로 도입 시기 조정", "impact_score": "3", "urgency_score": "2", "classification": "기회"},
                    {"id": "En3", "name": "탄소중립/RE100 요구 강화", "description": "중소기업의 실질적 이행 부담 가중", "impact_score": "3", "urgency_score": "3", "classification": "위협"}
                ],
                "opportunities": [{"rank": 1, "factor": "ESG 데이터 관리 수요", "action": "'원클릭 ESG 리포팅' 기능 개발"}],
                "threats": [{"rank": 1, "factor": "그린워싱 규제 강화", "mitigation": "데이터 근거 기반 진단 알고리즘 검증"}]
            },
            "legal": {
                "summary": "AI 규제는 높고 노동법은 엄격해지고 있어 컴플라이언스 대응이 시급합니다.",
                "issues": [
                    {"id": "L1", "name": "AI 기본법 및 데이터 규제", "description": "AI 신뢰성 확보 및 개인정보 보호 의무 강화 (2026 시행)", "impact_score": "5", "urgency_score": "5", "classification": "위협"},
                    {"id": "L2", "name": "사회적기업 인증 요건 완화", "description": "인증 진입장벽 완화로 잠재 고객군 확대", "impact_score": "3", "urgency_score": "3", "classification": "기회"},
                    {"id": "L3", "name": "노동법 처벌 강화", "description": "임금체불 및 근로시간 위반 제재 강화", "impact_score": "4", "urgency_score": "4", "classification": "위협"}
                ],
                "opportunities": [{"rank": 1, "factor": "규제 대응 솔루션 수요", "action": "자사 솔루션 '보안 인증' 취득"}],
                "threats": [{"rank": 1, "factor": "노무 리스크", "mitigation": "근로시간 관리 시스템 도입"}]
            }
        },
        "synthesis": {
            "top_5_opportunities": [
                {"rank": 1, "area": "Political", "factor": "정부의 AI·딥테크 육성 및 지역 균형 정책", "impact_score": "5", "action": "지역 특화 AI 사업 적극 수주"},
                {"rank": 2, "area": "Technological", "factor": "SLM(소형언어모델) 기술 효율성 증대", "impact_score": "5", "action": "고비용 LLM을 SLM으로 대체"},
                {"rank": 3, "area": "Environmental", "factor": "공급망 ESG 실사 확산", "impact_score": "4", "action": "중소기업용 ESG 대응 자동화 기능 출시"},
                {"rank": 4, "area": "Social", "factor": "지방 소멸 위기에 따른 에코시스템 빌더 역할", "impact_score": "5", "action": "지자체 연계 프로그램 확대"},
                {"rank": 5, "area": "Economic", "factor": "AI SaaS 시장 성장 및 바우처 지원", "impact_score": "4", "action": "정부 바우처 사업 공급기업 등록"}
            ],
            "top_5_threats": [
                {"rank": 1, "area": "Legal", "factor": "AI 기본법 및 개인정보보호 규제 강화", "urgency_score": "5", "mitigation": "데이터 비식별화 기술 적용"},
                {"rank": 2, "area": "Political", "factor": "높은 B2G 의존도 (80%)", "urgency_score": "5", "mitigation": "B2B SaaS 매출 비중 30%까지 확대"},
                {"rank": 3, "area": "Social", "factor": "내부 구성원 번아웃 및 과로", "urgency_score": "5", "mitigation": "자체 웰니스 솔루션 및 강제 휴식 제도"},
                {"rank": 4, "area": "Economic", "factor": "벤처투자 시장 침체", "urgency_score": "4", "mitigation": "매출 기반 현금 흐름 중시 경영"},
                {"rank": 5, "area": "Technological", "factor": "AI 기술 환각 및 신뢰성 문제", "urgency_score": "3", "mitigation": "RAG 기술 도입으로 정확도 향상"}
            ]
        }
    },
    
    "step_2_2_scenario": {
        "analysis_meta": {"company": "G임팩트"},
        "executive_summary": "G임팩트는 AI SaaS 시장 성장 기회와 높은 B2G 의존도/짧은 런웨이라는 약점이 공존합니다. '황금기' 시나리오보다 '지역의 봄' 또는 '수도권 독주' 시나리오 발생 가능성이 높습니다.",
        "uncertainty_analysis": {
            "axis_1": {"name": "정부 정책 기조", "positive_direction": "지원 확대", "negative_direction": "지원 축소"},
            "axis_2": {"name": "지역 경제 역동성", "positive_direction": "지역 분산", "negative_direction": "수도권 집중"}
        },
        "scenarios": {
            "scenario_1": {"quadrant": "++", "name": "황금기", "probability": "20%", "narrative": "정부의 전폭적인 지원과 지역 생태계 활황", "strategic_response": ["공격적 Scale-up", "시리즈 A 투자 유치"]},
            "scenario_2": {"quadrant": "-+", "name": "지역의 봄", "probability": "30%", "narrative": "정부 지원 축소, 민간 자생적 성장", "strategic_response": ["Pivot to Private", "SLM 도입 비용 절감"]},
            "scenario_3": {"quadrant": "--", "name": "빙하기", "probability": "15%", "narrative": "정부 지원 끊기고 지역 경제 무너짐", "strategic_response": ["비상 경영", "M&A Exit 모색"]},
            "scenario_4": {"quadrant": "+-", "name": "수도권 독주", "probability": "35%", "narrative": "기술 투자 증가하나 서울/판교에 집중", "strategic_response": ["Hybrid Operation", "온라인 SaaS로 전국화"]}
        },
        "robust_strategy": {
            "common_strategies": [
                "B2G 의존도 50% 이하로 축소",
                "SLM 도입으로 기술 비용 최적화",
                "지역을 초월한 투자자 네트워크 구축"
            ]
        }
    },
    
    "step_2_3_competition": {
        "analysis_meta": {"company": "G임팩트"},
        "five_forces": {
            "new_entrants": {"score": 3, "key_factors": ["낮은 제도적 진입장벽", "높은 신뢰 자본 장벽"]},
            "rivalry": {"score": 4.5, "key_factors": ["수도권 과밀", "공공기관과 서비스 중복"]},
            "substitutes": {"score": 2.5, "substitute_products": ["ChatGPT 등 범용 AI", "전통 경영 컨설팅"]},
            "supplier_power": {"score": 4, "key_suppliers": ["AI 개발자", "A급 창업 멘토"]},
            "buyer_power": {"score": 4, "customer_segments": ["B2G 80%", "B2B 스타트업"]},
            "overall": {"average_score": 3.6, "industry_attractiveness": "중간", "summary": "경쟁은 치열하나 'AI+지역+임팩트' 교집합은 블루오션"}
        },
        "competitor_analysis": {
            "business_competitors": [
                {"name": "SKT AI Lab", "type": "잠재경쟁", "strengths": ["압도적 기술 인프라", "SKT 사업 연계"], "weaknesses": ["지역 이해도 낮음"], "threat_level": "중간"},
                {"name": "이드로경영파트너스", "type": "직접경쟁", "strengths": ["전남 지역 기반", "공공사업 수주 경험"], "weaknesses": ["AI 기술 역량 부재"], "threat_level": "높음"},
                {"name": "구글 포 스타트업", "type": "간접경쟁", "strengths": ["글로벌 브랜드", "Google Cloud 지원"], "weaknesses": ["단기 프로그램", "사후 관리 부족"], "threat_level": "중간"},
                {"name": "광주청년창업 엑셀러레이팅", "type": "직접경쟁", "strengths": ["공공 신뢰", "직접 자금 지원"], "weaknesses": ["민간 투자 연결 부족"], "threat_level": "높음"}
            ],
            "impact_competitors": [
                {"name": "MYSC", "social_mission": "사회혁신 비즈니스 확산", "collaboration_potential": "높음", "collaboration_areas": ["JV 설립", "후속 투자 연계"]},
                {"name": "한국사회투자", "social_mission": "비즈니스로 더 좋은 세상", "collaboration_potential": "높음"},
                {"name": "임팩트스퀘어", "social_mission": "임팩트 비즈니스 가속화", "collaboration_potential": "중간"}
            ]
        }
    },
    
    "step_2_4_customer": {
        "analysis_meta": {"company": "G임팩트"},
        "customer_ecosystem": {
            "user": {
                "profile": "경영/투자/네트워크에 목마른 지역 기반의 초기 혁신가",
                "segments": ["지역 기술 스타트업", "예비창업자", "대기업 협력사(중소)", "사회적경제 조직"],
                "estimated_size": "약 8.6만 개사 (광주·전남)",
                "pain_points": ["자금 조달의 어려움", "고급 인력 채용 난항", "BM 고도화 역량 부족"],
                "jtbd": {"functional": ["투자 유치", "IR 자료 작성", "ESG 데이터 관리"], "emotional": ["사업 실패 불안감 해소"], "impact": ["지역 사회 기여"]}
            },
            "payer": {
                "profile": "예산 집행 효율성과 성과(KPI) 입증이 필요한 조직 리더",
                "types": ["B2G (지자체/공공기관)", "B2B (대기업 ESG팀)"],
                "decision_factors": ["정량적 성과 데이터", "운영 편의성/자동화", "비용 효율성"],
                "jtbd": {"functional": ["성과 리포팅 자동화", "피투자사 모니터링"], "emotional": ["감사 리스크 회피"], "impact": ["지역 불균형 해소"]}
            },
            "beneficiary": {
                "profile": "창업 활성화로 활력을 되찾는 지역 사회",
                "impact_metrics": ["지역 내 신규 고용 창출 수", "지역 자본 유치액"],
                "sdg_alignment": ["SDG 8 (일자리)", "SDG 10 (불평등 감소)"]
            }
        },
        "segment_priority_matrix": {
            "priority_1": {"segment": "지자체/공공기관 (B2G)", "reason": "검증된 예산, 가장 높은 접근성", "approach_strategy": "성과 관리 자동화 툴로 포지셔닝"},
            "priority_2": {"segment": "대기업 ESG팀 (B2B)", "reason": "높은 객단가, 반복 매출 가능성", "approach_strategy": "협력사 관리 비용 절감 솔루션"}
        }
    },
    
    "step_2_5_market": {
        "analysis_meta": {"company": "G임팩트"},
        "market_sizing": {
            "tam": {
                "triangulation": {"confirmed_tam": 120000},
                "top_down": {"value": "120000", "source": "중소벤처기업부"}
            },
            "sam": {
                "total": 9289,
                "by_customer_type": {"b2b": {"value": "7957"}, "b2g": {"value": "1332"}}
            },
            "som": {
                "year_1": {"value": 20, "sam_share": "0.2"},
                "year_3": {"value": 106, "sam_share": "1.1"},
                "year_5": {"value": 300, "sam_share": "3.2"}
            }
        },
        "market_trends": {
            "growth_rates": {
                "historical_cagr": {"value": "18", "period": "2021-2026"},
                "forecast_short": {"value": "5.2", "period": "1년"},
                "forecast_mid": {"value": "15", "period": "3년"}
            },
            "key_trends": [
                {"trend": "AI x Local", "impact": "지역 특화 데이터의 가치 상승", "opportunity_or_threat": "Opportunity"},
                {"trend": "ESG Supply Chain", "impact": "대기업의 협력사 관리 니즈 폭증", "opportunity_or_threat": "Opportunity"}
            ]
        }
    },
    
    "step_3_1_diagnosis": {
        "analysis_meta": {"company": "G임팩트"},
        "executive_summary": "G임팩트는 명확한 소셜 미션을 보유하나, 경영일반(2.6점)과 인사노무(3.0점) 영역에서 시스템 부재가 성장을 저해합니다. CEO를 포함한 핵심 인력의 과도한 업무 강도(휴일근무 6일)와 내부 피드백 시스템 부재(1점)는 심각한 번아웃을 초래할 위험이 큽니다.",
        "scores_summary": {
            "사회적가치": {"score": "3.1", "evaluation": "미션 명확, 협력 네트워크 부족"},
            "경영일반": {"score": "2.6", "evaluation": "가장 취약. 피드백 시스템 부재"},
            "영업마케팅": {"score": "3.3", "evaluation": "분석 우수, 실행력 보완 필요"},
            "재무": {"score": "4.0", "evaluation": "관리 양호, 정부 의존도 높음"},
            "인사조직": {"score": "3.0", "evaluation": "피로도 누적, 리더십 신뢰 부족"}
        },
        "improvement_priorities": [
            {"rank": 1, "area": "인사노무/경영일반", "issue": "구성원 번아웃 및 동기부여 시스템 붕괴", "urgency": "높음"},
            {"rank": 2, "area": "사회적가치", "issue": "외부 협력 네트워크 부재로 BM 실행력 저하", "urgency": "중간"},
            {"rank": 3, "area": "재무/판매", "issue": "짧은 런웨이 및 높은 정부 의존도", "urgency": "높음"}
        ]
    },
    
    "step_3_2_vrio": {
        "analysis_meta": {"company": "G임팩트"},
        "resource_identification": {
            "resources": [
                {"id": "R1", "name": "MYSC JV 파트너십", "type": "relational", "verification_status": "verified", "final_reliability": "✅"},
                {"id": "R2", "name": "광주·전남 로컬 네트워크", "type": "relational", "verification_status": "partially_verified", "final_reliability": "📊"},
                {"id": "R3", "name": "AI 엔진 4종", "type": "technological", "verification_status": "unverified", "final_reliability": "⚠️"},
                {"id": "R4", "name": "소셜 미션 및 브랜드", "type": "intangible", "verification_status": "verified", "final_reliability": "📊"},
                {"id": "R5", "name": "C-Level 인적 자원", "type": "human", "verification_status": "verified", "final_reliability": "📊"}
            ]
        },
        "vrio_evaluation": {
            "R1": {
                "resource_name": "MYSC JV 파트너십",
                "value": {"score": 5, "evidence": ["투자 연계 가능성", "브랜드 신뢰도 전이"]},
                "rarity": {"score": 5, "evidence": ["지역 AC 중 유일한 JV 사례"]},
                "imitability": {"score": 5, "evidence": ["장기간 신뢰 관계 필요"]},
                "organization": {"score": 2, "evidence": ["구체적 협력 프로세스 미비"]},
                "competitive_implication": "temporary"
            }
        },
        "portfolio_summary": {
            "sustained_advantage": [],
            "temporary_advantage": ["광주·전남 로컬 네트워크"],
            "competitive_parity": ["AI 엔진 4종", "소셜 미션 및 브랜드"],
            "competitive_disadvantage": ["C-Level 인적 자원 (번아웃 리스크)"]
        }
    },
    
    "step_3_3_swot": {
        "analysis_meta": {"company": "G임팩트"},
        "strengths": [
            {"id": "S1", "description": "MYSC JV 파트너십 (투자/네트워크 연계)", "impact_score": 5, "priority_score": 25, "reliability": "✅"},
            {"id": "S2", "description": "광주·전남 로컬 네트워크 (지역 진입장벽)", "impact_score": 4, "priority_score": 16, "reliability": "📊"},
            {"id": "S3", "description": "AI 엔진 4종 (업무 자동화/비용 효율)", "impact_score": 4, "priority_score": 12, "reliability": "📊"}
        ],
        "weaknesses": [
            {"id": "W1", "description": "조직 운영 시스템 부재 & 리더십 번아웃", "impact_score": 5, "priority_score": 20, "reliability": "📊"},
            {"id": "W2", "description": "높은 B2G 의존도 (80%) & 짧은 런웨이", "impact_score": 5, "priority_score": 15, "reliability": "📊"}
        ],
        "opportunities": [
            {"id": "O1", "description": "정부 지역소멸 대응/창업 예산 확대", "impact_score": 5, "probability_score": 5, "priority_score": 25, "reliability": "✅"},
            {"id": "O2", "description": "대기업 공급망 ESG 실사 의무화", "impact_score": 5, "probability_score": 4, "priority_score": 20, "reliability": "📊"}
        ],
        "threats": [
            {"id": "T1", "description": "벤처투자 시장 침체 (스타트업 구매력 저하)", "impact_score": 5, "probability_score": 5, "priority_score": 25, "reliability": "✅"},
            {"id": "T2", "description": "수도권 대형 AC의 지역 진출", "impact_score": 4, "probability_score": 4, "priority_score": 16, "reliability": "📊"}
        ],
        "key_insights": [
            "MYSC 파트너십과 정부 예산을 결합하여 단기 런웨이 확보가 급선무",
            "조직 운영 시스템 개선 없이는 어떤 전략도 지속 불가능",
            "대기업 ESG 시장 진입으로 B2G 의존도와 투자 침체를 동시 해결"
        ]
    },
    
    "step_3_4_tows": {
        "analysis_meta": {"company": "G임팩트"},
        "organization_capacity": {
            "total_members": 4,
            "workload_status": {"overtime_avg": 12.5, "holiday_work_avg": 4.5, "health_level": "위험"},
            "execution_capacity": {"available_bandwidth": "10%", "recommended_max_initiatives": 2}
        },
        "strategy_options": {
            "SO": [
                {"id": "SO-1", "name": "로컬 임팩트 메가 프로젝트", "hypothesis": "MYSC와 컨소시엄 구성으로 대형 지자체 사업 수주", "evaluation": {"total_score": 24, "priority": "⭐⭐⭐"}},
                {"id": "SO-2", "name": "AI ESG 공급망 솔루션", "hypothesis": "기존 AI 진단 엔진을 ESG용으로 패키징하여 B2B 진출", "evaluation": {"total_score": 21, "priority": "⭐⭐"}}
            ],
            "WO": [
                {"id": "WO-1", "name": "공공 자금 기반 조직 시스템화", "hypothesis": "정부 사업비로 운영 시스템 구축", "evaluation": {"total_score": 22, "priority": "⭐⭐⭐"}}
            ],
            "ST": [
                {"id": "ST-1", "name": "로컬 데이터 장벽 구축", "hypothesis": "지역 데이터 축적으로 진입장벽 형성", "evaluation": {"total_score": 21, "priority": "⭐⭐"}}
            ],
            "WT": [
                {"id": "WT-1", "name": "생존형 피봇 & 리더십 케어", "hypothesis": "불필요한 R&D 중단, 휴식", "evaluation": {"total_score": 22, "priority": "⭐⭐"}}
            ]
        },
        "strategy_sequencing": {
            "optimal_sequence": {
                "phase_1": {"period": "0-6개월", "strategies": ["WO-1", "WT-1"], "goals": "조직 안정화 및 번아웃 해소"},
                "phase_2": {"period": "6-12개월", "strategies": ["SO-1"], "goals": "대형 공공 사업 수주"},
                "phase_3": {"period": "1-2년", "strategies": ["SO-2", "ST-1"], "goals": "B2B SaaS 런칭"}
            }
        },
        "risk_management": {
            "pre_mortem": [
                {"failure_cause": "MYSC 협력이 MOU 수준에 그침", "probability": "중간", "preventive_action": "협력 전담자 지정 및 월 1회 정기 교류"},
                {"failure_cause": "리더십 번아웃으로 의사결정 지연", "probability": "높음", "preventive_action": "강제 휴가 및 R&R 재분배"},
                {"failure_cause": "B2B 전환 실패로 매출 정체", "probability": "중간", "preventive_action": "파일럿 고객 확보 후 확대"}
            ]
        },
        "decision_summary": {
            "top_3_strategies": [
                {"rank": 1, "id": "WO-1", "name": "공공 자금 기반 조직 시스템화", "type": "WO", "rationale": "조직 붕괴 리스크 해소 및 실행 기반 마련"},
                {"rank": 2, "id": "SO-1", "name": "로컬 임팩트 메가 프로젝트", "type": "SO", "rationale": "매출 증대 및 시장 지배력 확대"},
                {"rank": 3, "id": "SO-2", "name": "AI ESG 공급망 솔루션", "type": "SO", "rationale": "수익 다변화 및 스케일업"}
            ],
            "immediate_actions": [
                {"action": "CEO 주말 근무 중단 및 업무 위임 리스트 작성", "owner": "CEO", "deadline": "1주 내"},
                {"action": "MYSC 담당자와 공동 사업 기획 미팅", "owner": "CSO", "deadline": "2주 내"}
            ]
        }
    }
}

if __name__ == "__main__":
    import json
    with open('real_sample_data.json', 'w', encoding='utf-8') as f:
        json.dump(REAL_SAMPLE_DATA, f, ensure_ascii=False, indent=2)
    print("✅ 실제 데이터 기반 샘플 생성 완료")
