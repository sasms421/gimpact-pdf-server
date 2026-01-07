"""
G-IMPACT PDF Generation Server
FastAPI 기반 PDF 생성 API 서버

배포 옵션:
- Google Cloud Run
- Render / Railway
- AWS Lambda + API Gateway
"""

import os
import json
import base64
from io import BytesIO
from datetime import datetime
from typing import Optional, Dict, Any, List

from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

# ============================================
# FastAPI 앱 설정
# ============================================

app = FastAPI(
    title="G-IMPACT PDF Generator API",
    description="AI 분석 결과를 고품질 PDF 리포트로 변환하는 API",
    version="1.0.0"
)

# CORS 설정 (Google Apps Script에서 호출 허용)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ============================================
# 요청/응답 모델
# ============================================

class ReportMeta(BaseModel):
    business_name: str
    bm: Optional[str] = "ALL"
    collected_at: Optional[str] = None
    version: Optional[str] = "4.0"

class ReportOptions(BaseModel):
    generateSummary: bool = True
    generateDetail: bool = True
    businessName: str
    bm: Optional[str] = "ALL"

class TransformedData(BaseModel):
    sections: Dict[str, Any] = {}
    executiveSummary: Optional[str] = None

class GenerateRequest(BaseModel):
    meta: ReportMeta
    handoffs: Dict[str, Any]
    transformed: TransformedData
    options: ReportOptions

class GenerateResponse(BaseModel):
    success: bool
    summaryPdf: Optional[str] = None  # Base64 encoded
    detailPdf: Optional[str] = None   # Base64 encoded
    summaryPages: Optional[int] = None
    detailPages: Optional[int] = None
    error: Optional[str] = None
    generatedAt: Optional[str] = None

# ============================================
# API 엔드포인트
# ============================================

@app.get("/")
async def root():
    """헬스 체크"""
    return {
        "service": "G-IMPACT PDF Generator",
        "version": "1.0.0",
        "status": "healthy",
        "timestamp": datetime.now().isoformat()
    }

@app.get("/health")
async def health_check():
    """상세 헬스 체크"""
    font_ok = check_fonts()
    return {
        "status": "healthy" if font_ok else "degraded",
        "checks": {
            "pdf_generator": "ok",
            "fonts": "ok" if font_ok else "missing - will use fallback"
        },
        "timestamp": datetime.now().isoformat()
    }

@app.post("/generate", response_model=GenerateResponse)
async def generate_report(request: GenerateRequest):
    """
    PDF 리포트 생성
    
    - 요약 보고서 (15페이지, 디자인된 PDF)
    - 상세 보고서 (50-100페이지)
    
    Returns:
        Base64 인코딩된 PDF 데이터
    """
    try:
        result = GenerateResponse(
            success=True,
            generatedAt=datetime.now().isoformat()
        )
        
        # 데이터 준비
        report_data = prepare_report_data(request)
        
        # 요약 보고서 생성
        if request.options.generateSummary:
            summary_buf, summary_pages = generate_summary_report(
                report_data, 
                request.transformed,
                request.meta.business_name
            )
            result.summaryPdf = base64.b64encode(summary_buf.getvalue()).decode('utf-8')
            result.summaryPages = summary_pages
        
        # 상세 보고서 생성
        if request.options.generateDetail:
            detail_buf, detail_pages = generate_detail_report(
                report_data,
                request.transformed,
                request.meta.business_name
            )
            result.detailPdf = base64.b64encode(detail_buf.getvalue()).decode('utf-8')
            result.detailPages = detail_pages
        
        return result
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return GenerateResponse(
            success=False,
            error=str(e),
            generatedAt=datetime.now().isoformat()
        )

@app.post("/generate/summary")
async def generate_summary_only(request: GenerateRequest):
    """요약 보고서만 생성 (스트리밍 응답)"""
    try:
        report_data = prepare_report_data(request)
        pdf_buf, _ = generate_summary_report(
            report_data,
            request.transformed,
            request.meta.business_name
        )
        
        return StreamingResponse(
            pdf_buf,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename={request.meta.business_name}_요약보고서.pdf"
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/generate/detail")
async def generate_detail_only(request: GenerateRequest):
    """상세 보고서만 생성 (스트리밍 응답)"""
    try:
        report_data = prepare_report_data(request)
        pdf_buf, _ = generate_detail_report(
            report_data,
            request.transformed,
            request.meta.business_name
        )
        
        return StreamingResponse(
            pdf_buf,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename={request.meta.business_name}_상세보고서.pdf"
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ============================================
# 유틸리티 함수
# ============================================

def check_fonts() -> bool:
    """폰트 파일 존재 여부 확인"""
    font_paths = [
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "./fonts/NanumGothic.ttf",
        os.path.expanduser("~/.fonts/NanumGothic.ttf")
    ]
    return any(os.path.exists(p) for p in font_paths)

def prepare_report_data(request: GenerateRequest) -> Dict[str, Any]:
    """요청 데이터를 리포트 생성기 형식으로 변환"""
    data = {
        "company_name": request.meta.business_name,
        "bm": request.meta.bm,
        "generated_at": datetime.now().isoformat(),
    }
    
    # HANDOFF 데이터 직접 복사
    for key, value in request.handoffs.items():
        data[key] = value
    
    return data

# ============================================
# PDF 생성 로직
# ============================================

def generate_summary_report(
    data: Dict[str, Any], 
    transformed: TransformedData,
    company_name: str
) -> tuple[BytesIO, int]:
    """
    요약 보고서 생성 (15페이지)
    
    구조:
    - 표지 (1p)
    - 목차 (1p)
    - 1PAGE 요약 (1p)
    - 경영진 요약 (1-2p)
    - PESTEL 요약 (2p)
    - 시나리오 (1p)
    - 경쟁환경 (1p)
    - 고객/시장 (1p)
    - 경영진단 (1p)
    - VRIO (1p)
    - SWOT (1p)
    - TOWS 전략 (2p)
    """
    try:
        # 기존 PDF 생성기 임포트 시도
        from analysis_report_generator import GImpactReportGenerator
        
        generator = GImpactReportGenerator(data, company_name)
        pdf_buffer = generator.generate()
        
        # 페이지 수 계산 (대략적)
        pages = 15
        
        return pdf_buffer, pages
        
    except ImportError:
        # 폴백: 기본 PDF 생성
        return generate_basic_pdf(data, transformed, company_name, "summary")

def generate_detail_report(
    data: Dict[str, Any],
    transformed: TransformedData,
    company_name: str
) -> tuple[BytesIO, int]:
    """
    상세 보고서 생성 (50-100페이지)
    
    AI 변환된 텍스트를 사용하여 상세 보고서 생성
    """
    try:
        from detail_report_generator import DetailReportGenerator
        
        generator = DetailReportGenerator(data, transformed.dict(), company_name)
        pdf_buffer = generator.generate()
        pages = 50  # 대략적
        
        return pdf_buffer, pages
        
    except ImportError:
        # 폴백: 기본 PDF 생성
        return generate_basic_pdf(data, transformed, company_name, "detail")

def generate_basic_pdf(
    data: Dict[str, Any],
    transformed: TransformedData,
    company_name: str,
    report_type: str
) -> tuple[BytesIO, int]:
    """
    폴백용 기본 PDF 생성 (ReportLab 직접 사용)
    """
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    
    # 폰트 설정
    try:
        font_paths = [
            "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
            "./fonts/NanumGothic.ttf",
        ]
        for fp in font_paths:
            if os.path.exists(fp):
                pdfmetrics.registerFont(TTFont('NanumGothic', fp))
                break
    except:
        pass
    
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=20*mm,
        leftMargin=20*mm,
        topMargin=25*mm,
        bottomMargin=20*mm
    )
    
    styles = getSampleStyleSheet()
    
    # 커스텀 스타일
    try:
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Title'],
            fontName='NanumGothic',
            fontSize=24,
            textColor=colors.HexColor('#1a73e8'),
            spaceAfter=30
        )
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontName='NanumGothic',
            fontSize=10,
            leading=14
        )
    except:
        title_style = styles['Title']
        normal_style = styles['Normal']
    
    elements = []
    
    # 표지
    elements.append(Spacer(1, 100))
    elements.append(Paragraph("G-IMPACT 분석 리포트", title_style))
    elements.append(Spacer(1, 20))
    
    report_type_name = "요약 보고서" if report_type == "summary" else "상세 분석 보고서"
    elements.append(Paragraph(report_type_name, normal_style))
    elements.append(Spacer(1, 50))
    elements.append(Paragraph(f"<b>{company_name}</b>", normal_style))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(datetime.now().strftime("%Y년 %m월 %d일"), normal_style))
    
    elements.append(PageBreak())
    
    # 경영진 요약 (있으면)
    if transformed.executiveSummary:
        elements.append(Paragraph("경영진 요약", title_style))
        elements.append(Spacer(1, 10))
        
        # 마크다운 간단 변환
        lines = transformed.executiveSummary.split('\n')
        for line in lines:
            if line.strip():
                if line.startswith('## '):
                    elements.append(Paragraph(f"<b>{line[3:]}</b>", normal_style))
                elif line.startswith('### '):
                    elements.append(Paragraph(f"<b>{line[4:]}</b>", normal_style))
                else:
                    elements.append(Paragraph(line, normal_style))
                elements.append(Spacer(1, 5))
    
    # 상세 보고서인 경우 각 섹션 추가
    if report_type == "detail" and transformed.sections:
        for section_id, section_data in transformed.sections.items():
            elements.append(PageBreak())
            
            content = section_data.get('content', '') if isinstance(section_data, dict) else str(section_data)
            if content:
                lines = content.split('\n')
                for line in lines:
                    if line.strip():
                        if line.startswith('## '):
                            elements.append(Paragraph(f"<b>{line[3:]}</b>", title_style))
                        elif line.startswith('### '):
                            elements.append(Paragraph(f"<b>{line[4:]}</b>", normal_style))
                        else:
                            elements.append(Paragraph(line, normal_style))
                        elements.append(Spacer(1, 3))
    
    doc.build(elements)
    buffer.seek(0)
    
    pages = 15 if report_type == "summary" else 50
    return buffer, pages

# ============================================
# 로컬 실행
# ============================================

if __name__ == "__main__":
    import uvicorn
    
    port = int(os.environ.get("PORT", 8080))
    uvicorn.run(
        "pdf_api_server:app",
        host="0.0.0.0",
        port=port,
        reload=True
    )
