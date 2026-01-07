#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
G-IMPACT 분석 리포트 생성기 v2.0
Analysis Report Generator (2.1~3.4 단계)

구조:
1. 1PAGE 요약 (1페이지)
2. 경영진용 요약 (2-3페이지)  
3. 단계별 상세 리포트 (2.1~3.4)
"""

import json
import os
from datetime import datetime
from io import BytesIO

# ReportLab imports
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, 
    PageBreak, Image, KeepTogether, HRFlowable, ListFlowable, ListItem
)
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping

# Matplotlib for charts
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

# ==============================================================================
# 색상 테마
# ==============================================================================
COLORS = {
    'primary': colors.HexColor('#2563EB'),     # 파랑 (메인)
    'secondary': colors.HexColor('#1E40AF'),   # 진파랑
    'accent': colors.HexColor('#10B981'),      # 초록 (강점/기회)
    'warning': colors.HexColor('#F59E0B'),     # 주황 (주의)
    'danger': colors.HexColor('#EF4444'),      # 빨강 (위협/약점)
    'dark': colors.HexColor('#1F2937'),
    'medium': colors.HexColor('#6B7280'),
    'gray': colors.HexColor('#9CA3AF'),
    'light': colors.HexColor('#F3F4F6'),
    'white': colors.white,
    # 분석 특화 색상
    'opportunity': colors.HexColor('#10B981'), # 기회 = 초록
    'threat': colors.HexColor('#EF4444'),      # 위협 = 빨강
    'strength': colors.HexColor('#10B981'),    # 강점 = 초록
    'weakness': colors.HexColor('#EF4444'),    # 약점 = 빨강
    # 섹션별 색상
    'pestel': colors.HexColor('#7C3AED'),      # 보라 (PESTEL)
    'scenario': colors.HexColor('#0891B2'),    # 청록 (시나리오)
    'competition': colors.HexColor('#EA580C'), # 주황 (경쟁)
    'customer': colors.HexColor('#0D9488'),    # 틸 (고객)
    'market': colors.HexColor('#2563EB'),      # 파랑 (시장)
    'diagnosis': colors.HexColor('#4F46E5'),   # 인디고 (경영진단)
    'vrio': colors.HexColor('#7C3AED'),        # 보라 (VRIO)
    'swot': colors.HexColor('#059669'),        # 에메랄드 (SWOT)
    'tows': colors.HexColor('#1E40AF'),        # 네이비 (TOWS)
}

# ==============================================================================
# 폰트 설정
# ==============================================================================
def setup_fonts():
    """한글 폰트 설정"""
    font_path = '/usr/share/fonts/truetype/nanum/NanumGothic.ttf'
    bold_path = '/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf'
    
    try:
        pdfmetrics.registerFont(TTFont('NanumGothic', font_path))
        pdfmetrics.registerFont(TTFont('NanumGothicBold', bold_path))
        addMapping('NanumGothic', 0, 0, 'NanumGothic')
        addMapping('NanumGothic', 1, 0, 'NanumGothicBold')
        addMapping('NanumGothic', 0, 1, 'NanumGothic')
        addMapping('NanumGothic', 1, 1, 'NanumGothicBold')
    except Exception as e:
        print(f"폰트 등록 오류: {e}")
    
    # Matplotlib 폰트
    fm.fontManager.addfont(font_path)
    prop = fm.FontProperties(fname=font_path)
    plt.rcParams['font.family'] = prop.get_name()
    plt.rcParams['axes.unicode_minus'] = False

setup_fonts()

FONT = 'NanumGothic'
FONT_BOLD = 'NanumGothicBold'

# ==============================================================================
# 스타일 정의
# ==============================================================================
def create_styles():
    styles = getSampleStyleSheet()
    
    styles.add(ParagraphStyle('KTitle', fontName=FONT_BOLD, fontSize=24, leading=30,
                              alignment=TA_CENTER, textColor=COLORS['dark'], spaceAfter=20))
    styles.add(ParagraphStyle('KH1', fontName=FONT_BOLD, fontSize=16, leading=22,
                              textColor=COLORS['primary'], spaceBefore=15, spaceAfter=10))
    styles.add(ParagraphStyle('KH2', fontName=FONT_BOLD, fontSize=13, leading=18,
                              textColor=COLORS['secondary'], spaceBefore=12, spaceAfter=8))
    styles.add(ParagraphStyle('KH3', fontName=FONT_BOLD, fontSize=11, leading=15,
                              textColor=COLORS['dark'], spaceBefore=8, spaceAfter=5))
    styles.add(ParagraphStyle('KBody', fontName=FONT, fontSize=10, leading=14,
                              textColor=COLORS['dark'], alignment=TA_JUSTIFY, spaceAfter=6))
    styles.add(ParagraphStyle('KBodySmall', fontName=FONT, fontSize=9, leading=12,
                              textColor=COLORS['medium'], spaceAfter=4))
    styles.add(ParagraphStyle('KCaption', fontName=FONT, fontSize=8, leading=10,
                              textColor=COLORS['medium'], alignment=TA_CENTER, spaceAfter=8))
    styles.add(ParagraphStyle('KBullet', fontName=FONT, fontSize=10, leading=14,
                              textColor=COLORS['dark'], leftIndent=15, spaceAfter=3))
    return styles

# ==============================================================================
# 차트 생성 함수
# ==============================================================================
def create_horizontal_bar_chart(data, labels, title, max_val=5, width=400, height=220):
    """수평 막대 차트"""
    fig, ax = plt.subplots(figsize=(width/100, height/100), dpi=100)
    
    colors_list = ['#10B981' if v >= 4 else '#3B82F6' if v >= 3 else '#F59E0B' if v >= 2 else '#EF4444'
                   for v in data]
    
    y_pos = np.arange(len(labels))
    bars = ax.barh(y_pos, data, color=colors_list, height=0.6)
    
    for bar, val in zip(bars, data):
        ax.annotate(f'{val:.1f}', xy=(bar.get_width() + 0.1, bar.get_y() + bar.get_height()/2),
                    ha='left', va='center', fontsize=10, fontweight='bold')
    
    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels, fontsize=10)
    ax.set_xlim(0, max_val + 0.8)
    ax.set_title(title, fontsize=12, fontweight='bold', pad=10)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.invert_yaxis()
    ax.xaxis.grid(True, linestyle='--', alpha=0.3)
    
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    buf.seek(0)
    return buf

def create_diagnosis_radar_only(scores_dict, width=280, height=280):
    """레이더 차트만 생성 (테이블은 reportlab으로 별도 생성)"""
    
    fig, ax = plt.subplots(figsize=(width/100, height/100), subplot_kw=dict(polar=True), dpi=100)
    
    labels = list(scores_dict.keys())
    values = [float(scores_dict[k].get('score', 0)) for k in labels]
    
    # 짧은 라벨
    short_labels = []
    for l in labels:
        if '사회적' in l:
            short_labels.append('사회적\n가치')
        elif '영업' in l:
            short_labels.append('영업\n마케팅')
        elif '경영' in l:
            short_labels.append('경영\n일반')
        elif '인사' in l:
            short_labels.append('인사\n조직')
        else:
            short_labels.append(l)
    
    N = len(labels)
    angles = [n / float(N) * 2 * np.pi for n in range(N)]
    angles += angles[:1]
    values_plot = values + [values[0]]
    
    # 배경
    for i in range(1, 6):
        ax.plot(angles, [i] * (N + 1), color='#E5E7EB', linewidth=0.5, linestyle='--')
    
    # 데이터
    ax.fill(angles, values_plot, color='#3B82F6', alpha=0.25)
    ax.plot(angles, values_plot, color='#2563EB', linewidth=2)
    ax.scatter(angles[:-1], values, color='#1E40AF', s=50, zorder=5)
    
    # 점수 값 표시
    for angle, val in zip(angles[:-1], values):
        # 값 위치 조정 (바깥쪽으로)
        r_offset = val + 0.5 if val < 4 else val - 0.5
        ax.text(angle, r_offset, f'{val:.1f}', ha='center', va='center',
               fontsize=9, fontweight='bold', color='#1E40AF')
    
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(short_labels, fontsize=9)
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(['1', '2', '3', '4', '5'], fontsize=7, color='#9CA3AF')
    ax.spines['polar'].set_color('#E5E7EB')
    
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    buf.seek(0)
    return buf

def create_score_horizontal_bar(scores_dict, width=380, height=140):
    """수평 막대 점수 차트 - 1PAGE 요약용"""
    fig, ax = plt.subplots(figsize=(width/100, height/100), dpi=100)
    
    # 데이터 준비
    labels = list(scores_dict.keys())
    values = [float(scores_dict[k].get('score', 0)) for k in labels]
    
    # 짧은 라벨
    short_labels = []
    for l in labels:
        if '사회적' in l:
            short_labels.append('사회적가치')
        elif '영업' in l:
            short_labels.append('영업마케팅')
        elif '경영' in l:
            short_labels.append('경영일반')
        elif '인사' in l:
            short_labels.append('인사조직')
        else:
            short_labels.append(l)
    
    # 색상 (점수에 따라)
    bar_colors = []
    for v in values:
        if v >= 4:
            bar_colors.append('#10B981')  # 초록 (양호)
        elif v >= 3:
            bar_colors.append('#F59E0B')  # 주황 (보통)
        else:
            bar_colors.append('#EF4444')  # 빨강 (취약)
    
    y_pos = np.arange(len(labels))
    
    # 배경 그리드
    for i in range(1, 6):
        ax.axvline(x=i, color='#E5E7EB', linewidth=0.5, linestyle='--', zorder=0)
    
    # 막대 그리기
    bars = ax.barh(y_pos, values, height=0.6, color=bar_colors, edgecolor='white', linewidth=1)
    
    # 점수 라벨
    for i, (bar, v) in enumerate(zip(bars, values)):
        # 막대 끝에 점수 표시
        ax.text(v + 0.15, bar.get_y() + bar.get_height()/2, f'{v:.1f}', 
                va='center', ha='left', fontsize=10, fontweight='bold', color='#1F2937')
        
        # 상태 텍스트
        if v >= 4:
            status = '양호'
            status_color = '#059669'
        elif v >= 3:
            status = '보통'
            status_color = '#D97706'
        else:
            status = '취약'
            status_color = '#DC2626'
        ax.text(5.3, bar.get_y() + bar.get_height()/2, status, 
                va='center', ha='left', fontsize=8, fontweight='bold', color=status_color)
    
    ax.set_yticks(y_pos)
    ax.set_yticklabels(short_labels, fontsize=9)
    ax.set_xlim(0, 5.8)
    ax.set_xticks([1, 2, 3, 4, 5])
    ax.set_xticklabels(['1', '2', '3', '4', '5'], fontsize=8)
    ax.invert_yaxis()
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#E5E7EB')
    ax.spines['bottom'].set_color('#E5E7EB')
    
    # 범례
    ax.text(5.8, -0.5, '■ 양호(4+)  ■ 보통(3+)  ■ 취약(<3)', 
            ha='right', fontsize=7, color='#6B7280')
    
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white', pad_inches=0.05)
    plt.close()
    buf.seek(0)
    return buf

# 호환성을 위해 기존 함수명 유지
def create_diagnosis_combo_chart(scores_dict, width=280, height=280):
    """레이더 차트 생성 (테이블은 별도)"""
    return create_diagnosis_radar_only(scores_dict, width, height)

def create_concentric_market_chart(tam, sam, som, width=350, height=350):
    """동심원 버블 차트 - 시장 규모 (완전한 정원 보장)"""
    # 정사각형 figure 생성
    fig = plt.figure(figsize=(5, 5), dpi=100)
    ax = fig.add_axes([0.05, 0.05, 0.9, 0.9])  # 정사각형 axes
    
    # 원을 왼쪽으로, 범례를 오른쪽으로
    cx, cy = 0.30, 0.50
    max_r = 0.24
    
    tam_r = max_r
    sam_r = max_r * 0.55
    som_r = max_r * 0.22
    
    # 동심원
    circle1 = plt.Circle((cx, cy), tam_r, color='#DBEAFE', ec='#3B82F6', linewidth=2.5)
    circle2 = plt.Circle((cx, cy), sam_r, color='#93C5FD', ec='#2563EB', linewidth=2.5)
    circle3 = plt.Circle((cx, cy), som_r, color='#2563EB', ec='#1E40AF', linewidth=2.5)
    ax.add_patch(circle1)
    ax.add_patch(circle2)
    ax.add_patch(circle3)
    
    # 원 내부 라벨
    ax.text(cx, cy + tam_r - 0.035, 'TAM', fontsize=9, fontweight='bold', ha='center', color='#1E40AF')
    ax.text(cx, cy + tam_r - 0.07, f'{tam:,.0f}억', fontsize=8, ha='center', color='#3B82F6')
    ax.text(cx, cy + sam_r - 0.025, 'SAM', fontsize=8, fontweight='bold', ha='center', color='#1E40AF')
    ax.text(cx, cy + sam_r - 0.055, f'{sam:,.0f}억', fontsize=7, ha='center', color='#2563EB')
    ax.text(cx, cy + 0.015, 'SOM', fontsize=8, fontweight='bold', ha='center', color='white')
    ax.text(cx, cy - 0.015, f'{som:,.0f}억', fontsize=7, ha='center', color='white')
    
    # 우측 범례
    legend_x = 0.68
    descriptions = [
        ('TAM', f'{tam:,.0f}억', '전체 시장 규모', '#DBEAFE', '#3B82F6'),
        ('SAM', f'{sam:,.0f}억', '접근 가능 시장', '#93C5FD', '#2563EB'),
        ('SOM', f'{som:,.0f}억', '1년차 획득 목표', '#2563EB', '#1E40AF'),
    ]
    
    for i, (name, value, desc, bg_color, text_color) in enumerate(descriptions):
        y = 0.78 - i * 0.24
        legend_circle = plt.Circle((legend_x, y), 0.022, color=bg_color, ec=text_color, linewidth=1.5)
        ax.add_patch(legend_circle)
        text_x = legend_x + 0.05
        ax.text(text_x, y + 0.03, name, fontsize=10, fontweight='bold', color='#1F2937')
        ax.text(text_x, y - 0.005, value, fontsize=11, fontweight='bold', color=text_color)
        ax.text(text_x, y - 0.045, desc, fontsize=7, color='#6B7280')
    
    ax.set_title('시장 규모 (TAM → SAM → SOM)', fontsize=11, fontweight='bold', pad=8, color='#1F2937')
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.set_aspect('equal')  # 핵심: 정원 보장
    ax.axis('off')
    
    buf = BytesIO()
    # bbox_inches='tight' 제거하여 aspect ratio 유지
    plt.savefig(buf, format='png', dpi=150, facecolor='white', 
                bbox_inches=None, pad_inches=0)
    plt.close()
    buf.seek(0)
    return buf

def create_radar_chart(categories, values, title, max_val=5, width=320, height=320):
    """레이더 차트"""
    N = len(categories)
    angles = [n / float(N) * 2 * np.pi for n in range(N)]
    angles += angles[:1]
    values = list(values) + [values[0]]
    
    fig, ax = plt.subplots(figsize=(width/100, height/100), subplot_kw=dict(polar=True), dpi=100)
    ax.plot(angles, values, 'o-', linewidth=2, color='#2563EB')
    ax.fill(angles, values, alpha=0.25, color='#2563EB')
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=9)
    ax.set_ylim(0, max_val)
    ax.set_title(title, fontsize=12, fontweight='bold', pad=15)
    
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    buf.seek(0)
    return buf

def create_scenario_matrix(scenarios, width=400, height=320):
    """시나리오 2x2 매트릭스"""
    fig, ax = plt.subplots(figsize=(width/100, height/100), dpi=100)
    
    # 배경 사분면
    ax.fill_between([0, 1], 0, 1, alpha=0.15, color='#10B981')   # ++
    ax.fill_between([-1, 0], 0, 1, alpha=0.15, color='#F59E0B')  # -+
    ax.fill_between([-1, 0], -1, 0, alpha=0.15, color='#EF4444') # --
    ax.fill_between([0, 1], -1, 0, alpha=0.15, color='#3B82F6')  # +-
    
    ax.axhline(y=0, color='gray', linewidth=1)
    ax.axvline(x=0, color='gray', linewidth=1)
    
    positions = {'++': (0.5, 0.5), '-+': (-0.5, 0.5), '--': (-0.5, -0.5), '+-': (0.5, -0.5)}
    
    for s in scenarios:
        quadrant = s.get('quadrant', '++')
        name = s.get('name', '')
        prob = s.get('probability', '')
        if quadrant in positions:
            x, y = positions[quadrant]
            ax.scatter(x, y, s=300, c='#1E40AF', zorder=5, edgecolors='white', linewidth=2)
            ax.annotate(f"{name}\n({prob})", xy=(x, y), xytext=(0, -35),
                       textcoords='offset points', ha='center', fontsize=9, fontweight='bold')
    
    ax.set_xlim(-1.1, 1.1)
    ax.set_ylim(-1.1, 1.1)
    ax.set_xlabel('정부 정책 기조 →', fontsize=10)
    ax.set_ylabel('지역 경제 역동성 →', fontsize=10)
    ax.set_title('시나리오 매트릭스', fontsize=12, fontweight='bold')
    
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    buf.seek(0)
    return buf

def create_scenario_probability_chart(scenarios, width=280, height=200):
    """시나리오 확률 도넛 차트"""
    fig, ax = plt.subplots(figsize=(width/100, height/100), dpi=100)
    
    names = []
    probs = []
    colors_list = ['#10B981', '#F59E0B', '#EF4444', '#3B82F6']  # ++, -+, --, +-
    
    quadrant_order = ['++', '-+', '--', '+-']
    color_map = {q: c for q, c in zip(quadrant_order, colors_list)}
    
    for s in scenarios:
        names.append(s.get('name', '')[:8])
        prob_str = s.get('probability', '0%').replace('%', '')
        try:
            probs.append(float(prob_str))
        except:
            probs.append(0)
    
    # 색상 매핑
    chart_colors = [color_map.get(s.get('quadrant', '++'), '#6B7280') for s in scenarios]
    
    # 도넛 차트
    wedges, texts, autotexts = ax.pie(probs, labels=names, colors=chart_colors,
                                       autopct='%1.0f%%', startangle=90,
                                       wedgeprops=dict(width=0.5),
                                       textprops={'fontsize': 8})
    
    for autotext in autotexts:
        autotext.set_fontsize(9)
        autotext.set_fontweight('bold')
    
    ax.set_title('시나리오 확률 분포', fontsize=11, fontweight='bold', pad=10)
    
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    buf.seek(0)
    return buf

def create_strategy_roadmap(strategies, width=480, height=180):
    """전략 로드맵 - 간트 차트 스타일"""
    fig, ax = plt.subplots(figsize=(width/80, height/80), dpi=120)
    
    ax.set_facecolor('white')
    
    # 전략 색상
    strategy_colors = {
        'WO': '#3B82F6',   # 파랑 (전환)
        'SO': '#10B981',   # 초록 (공격)
        'ST': '#F59E0B',   # 주황 (방어)
        'WT': '#EF4444',   # 빨강 (생존)
    }
    
    # 전략별 시작/종료 기간 (순위에 따라)
    strategy_schedule = [
        {'start': 0, 'end': 6, 'phase': 'Phase 1 (0-6M)'},    # 1순위: 즉시 시작
        {'start': 3, 'end': 12, 'phase': 'Phase 1-2 (3-12M)'}, # 2순위: 3개월 후 시작
        {'start': 6, 'end': 24, 'phase': 'Phase 2-3 (6-24M)'}, # 3순위: 6개월 후 시작
    ]
    
    y_positions = [2.5, 1.5, 0.5]
    bar_height = 0.6
    
    # Phase 구분선
    for month in [6, 12]:
        ax.axvline(x=month, color='#E5E7EB', linewidth=1, linestyle='--', zorder=0)
    
    # Phase 라벨 (상단)
    ax.text(3, 3.3, 'Phase 1', ha='center', fontsize=9, fontweight='bold', color='#3B82F6')
    ax.text(3, 3.0, '조직 안정화', ha='center', fontsize=7, color='#6B7280')
    ax.text(9, 3.3, 'Phase 2', ha='center', fontsize=9, fontweight='bold', color='#10B981')
    ax.text(9, 3.0, '사업 확장', ha='center', fontsize=7, color='#6B7280')
    ax.text(18, 3.3, 'Phase 3', ha='center', fontsize=9, fontweight='bold', color='#F59E0B')
    ax.text(18, 3.0, '스케일업', ha='center', fontsize=7, color='#6B7280')
    
    for i, s in enumerate(strategies[:3]):
        name = s.get('name', '')
        stype = s.get('type', 'SO')
        rank = s.get('rank', i + 1)
        color = strategy_colors.get(stype, '#6B7280')
        
        schedule = strategy_schedule[i]
        start = schedule['start']
        end = schedule['end']
        duration = end - start
        y = y_positions[i]
        
        # 막대 그리기 (그림자 효과)
        shadow = plt.Rectangle((start + 0.1, y - bar_height/2 - 0.05), duration, bar_height,
                               color='#00000015', zorder=1)
        ax.add_patch(shadow)
        
        # 메인 막대
        from matplotlib.patches import FancyBboxPatch
        bar = FancyBboxPatch((start, y - bar_height/2), duration, bar_height,
                             boxstyle="round,pad=0.02,rounding_size=0.08",
                             facecolor=color, edgecolor='white', linewidth=2, zorder=2)
        ax.add_patch(bar)
        
        # 순위 원형 배지 (막대 시작점)
        badge = plt.Circle((start + 0.5, y), 0.25, facecolor='white', 
                          edgecolor=color, linewidth=2, zorder=3)
        ax.add_patch(badge)
        ax.text(start + 0.5, y, str(rank), ha='center', va='center',
               fontsize=10, fontweight='bold', color=color, zorder=4)
        
        # 전략명 (막대 중앙)
        display_name = name[:12] + '..' if len(name) > 12 else name
        ax.text(start + duration/2 + 0.3, y, display_name, ha='center', va='center',
               fontsize=8, fontweight='bold', color='white', zorder=4)
        
        # 전략 유형 + 기간 (막대 오른쪽)
        ax.text(end + 0.3, y, f'{stype}', ha='left', va='center',
               fontsize=8, fontweight='bold', color=color, zorder=4)
        ax.text(end + 0.3, y - 0.25, f'{start}-{end}M', ha='left', va='center',
               fontsize=7, color='#6B7280', zorder=4)
    
    # X축 (시간)
    ax.set_xlim(-0.5, 26)
    ax.set_ylim(-0.2, 3.6)
    ax.set_xticks([0, 6, 12, 18, 24])
    ax.set_xticklabels(['현재', '6M', '12M', '18M', '24M'], fontsize=8)
    ax.set_yticks([])
    
    # 테두리 정리
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_color('#E5E7EB')
    
    # 범례
    legend_y = -0.1
    legend_items = [('SO 공격', '#10B981'), ('WO 전환', '#3B82F6'), 
                    ('ST 방어', '#F59E0B'), ('WT 생존', '#EF4444')]
    for j, (label, lcolor) in enumerate(legend_items):
        ax.add_patch(plt.Rectangle((j*5.5 + 1, legend_y - 0.15), 0.8, 0.3, 
                                   facecolor=lcolor, zorder=5))
        ax.text(j*5.5 + 2, legend_y, label, fontsize=7, va='center', color='#4B5563')
    
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white', pad_inches=0.05)
    plt.close()
    buf.seek(0)
    return buf

def create_five_forces_chart(forces_data, width=400, height=300):
    """Five Forces 차트 - 라벨 개선"""
    labels = ['신규진입', '경쟁강도', '대체재', '공급자', '구매자']
    values = [
        forces_data.get('new_entrants', {}).get('score', 0),
        forces_data.get('rivalry', {}).get('score', 0),
        forces_data.get('substitutes', {}).get('score', 0),
        forces_data.get('supplier_power', {}).get('score', 0),
        forces_data.get('buyer_power', {}).get('score', 0),
    ]
    
    return create_radar_chart(labels, values, 'Five Forces 분석', max_val=5, width=width, height=height)

def create_market_funnel(tam, sam, som, width=350, height=250):
    """시장 규모 퍼널 차트"""
    fig, ax = plt.subplots(figsize=(width/100, height/100), dpi=100)
    
    # 퍼널 데이터
    data = [tam, sam, som]
    labels = [f'TAM\n{tam:,.0f}억', f'SAM\n{sam:,.0f}억', f'SOM\n{som:,.0f}억']
    colors_list = ['#93C5FD', '#3B82F6', '#1E40AF']
    
    # 가로 막대로 퍼널 표현
    y_pos = [2, 1, 0]
    widths = [d / tam for d in data]
    
    for i, (y, w, label, c) in enumerate(zip(y_pos, widths, labels, colors_list)):
        ax.barh(y, w, height=0.7, color=c, left=(1-w)/2)
        ax.text(0.5, y, label, ha='center', va='center', fontsize=10, fontweight='bold', color='white')
    
    ax.set_xlim(0, 1)
    ax.set_ylim(-0.5, 2.5)
    ax.axis('off')
    ax.set_title('시장 규모 (TAM → SAM → SOM)', fontsize=12, fontweight='bold')
    
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    buf.seek(0)
    return buf

# ==============================================================================
# 테이블 생성 함수
# ==============================================================================
def styled_table(data, col_widths=None, header_color=None):
    """스타일 테이블"""
    if header_color is None:
        header_color = COLORS['primary']
    
    table = Table(data, colWidths=col_widths)
    style = [
        ('BACKGROUND', (0, 0), (-1, 0), header_color),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('FONTNAME', (0, 1), (-1, -1), FONT),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, COLORS['light']),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]
    # 교대 행 색상
    for i in range(1, len(data)):
        if i % 2 == 0:
            style.append(('BACKGROUND', (0, i), (-1, i), COLORS['light']))
    
    table.setStyle(TableStyle(style))
    return table

# ==============================================================================
# 페이지 템플릿
# ==============================================================================
class ReportTemplate:
    def __init__(self, company_name, report_date):
        self.company_name = company_name
        self.report_date = report_date
        self.page_num = 0
    
    def cover_page(self, canvas, doc):
        """표지 - 프리미엄 디자인"""
        canvas.saveState()
        w, h = A4
        
        # 배경 그라데이션 효과 (여러 층의 사각형)
        gradient_colors = [
            (0, '#1E40AF'),    # 진한 파랑 (상단)
            (0.3, '#2563EB'),  # 메인 파랑
            (0.6, '#3B82F6'),  # 밝은 파랑
        ]
        
        # 상단 메인 배경 (그라데이션 효과)
        canvas.setFillColor(colors.HexColor('#1E40AF'))
        canvas.rect(0, h - 145*mm, w, 145*mm, fill=True, stroke=False)
        
        # 장식 요소 - 대각선 스트라이프
        canvas.setStrokeColor(colors.HexColor('#3B82F6'))
        canvas.setLineWidth(0.5)
        for i in range(10):
            y_offset = h - 30*mm - i * 12*mm
            canvas.line(0, y_offset, w, y_offset + 30*mm)
        
        # 상단 악센트 바
        canvas.setFillColor(colors.HexColor('#10B981'))
        canvas.rect(0, h - 8*mm, w, 8*mm, fill=True, stroke=False)
        
        # 메인 제목 영역
        canvas.setFillColor(colors.white)
        canvas.setFont(FONT_BOLD, 38)
        canvas.drawCentredString(w/2, h - 55*mm, "분석 리포트")
        
        # 부제목
        canvas.setFont(FONT, 14)
        canvas.setFillColor(colors.HexColor('#93C5FD'))
        canvas.drawCentredString(w/2, h - 72*mm, "G-IMPACT Analysis Report")
        
        # 구분선
        canvas.setStrokeColor(colors.HexColor('#60A5FA'))
        canvas.setLineWidth(2)
        canvas.line(w/2 - 60*mm, h - 85*mm, w/2 + 60*mm, h - 85*mm)
        
        # 회사명 (강조 박스)
        canvas.setFillColor(colors.white)
        canvas.roundRect(w/2 - 70*mm, h - 135*mm, 140*mm, 35*mm, 8, fill=True, stroke=False)
        
        canvas.setFillColor(COLORS['primary'])
        canvas.setFont(FONT_BOLD, 30)
        canvas.drawCentredString(w/2, h - 122*mm, self.company_name)
        
        # 하단 메타 정보 영역
        canvas.setFillColor(colors.HexColor('#F8FAFC'))
        canvas.roundRect(35*mm, 28*mm, w - 70*mm, 55*mm, 8, fill=True, stroke=False)
        
        # 메타 정보 테두리
        canvas.setStrokeColor(colors.HexColor('#E2E8F0'))
        canvas.setLineWidth(1)
        canvas.roundRect(35*mm, 28*mm, w - 70*mm, 55*mm, 8, fill=False, stroke=True)
        
        # 메타 정보 텍스트
        canvas.setFillColor(COLORS['dark'])
        canvas.setFont(FONT_BOLD, 10)
        canvas.drawCentredString(w/2, 72*mm, "리포트 정보")
        
        canvas.setFont(FONT, 10)
        canvas.setFillColor(COLORS['medium'])
        canvas.drawCentredString(w/2, 58*mm, f"생성일: {self.report_date}")
        canvas.drawCentredString(w/2, 46*mm, "분석 범위: 2.1 PESTEL ~ 3.4 TOWS")
        canvas.drawCentredString(w/2, 34*mm, "버전: 3.0")
        
        # 하단 브랜딩 바
        canvas.setFillColor(colors.HexColor('#1E40AF'))
        canvas.rect(0, 0, w, 12*mm, fill=True, stroke=False)
        canvas.setFillColor(colors.white)
        canvas.setFont(FONT, 8)
        canvas.drawCentredString(w/2, 4*mm, "Powered by G-IMPACT Analysis Engine")
        
        canvas.restoreState()
    
    def first_content_page(self, canvas, doc):
        """첫 번째 콘텐츠 페이지 (표지 다음) - 헤더/푸터만"""
        self.header_footer(canvas, doc)
    
    def header_footer(self, canvas, doc):
        """헤더/푸터"""
        canvas.saveState()
        w, h = A4
        
        # 헤더
        canvas.setFillColor(COLORS['primary'])
        canvas.rect(0, h - 18*mm, w, 18*mm, fill=True, stroke=False)
        
        canvas.setFillColor(colors.white)
        canvas.setFont(FONT_BOLD, 10)
        canvas.drawString(15*mm, h - 12*mm, f"{self.company_name} 분석 리포트")
        canvas.setFont(FONT, 9)
        canvas.drawRightString(w - 15*mm, h - 12*mm, "G-IMPACT Analysis Report")
        
        # 푸터
        self.page_num += 1
        canvas.setFillColor(COLORS['medium'])
        canvas.setFont(FONT, 8)
        canvas.drawCentredString(w/2, 10*mm, f"- {self.page_num} -")
        canvas.drawString(15*mm, 10*mm, self.report_date)
        
        canvas.restoreState()

# ==============================================================================
# 리포트 빌더
# ==============================================================================
class AnalysisReportBuilder:
    def __init__(self, data, company_name):
        self.data = data
        self.company_name = company_name
        self.styles = create_styles()
        self.elements = []
    
    def add_h1(self, text):
        self.elements.append(Paragraph(text, self.styles['KH1']))
    
    def add_h2(self, text):
        self.elements.append(Paragraph(text, self.styles['KH2']))
    
    def add_h3(self, text):
        self.elements.append(Paragraph(text, self.styles['KH3']))
    
    def add_body(self, text):
        self.elements.append(Paragraph(text, self.styles['KBody']))
    
    def add_small(self, text):
        self.elements.append(Paragraph(text, self.styles['KBodySmall']))
    
    def add_bullet(self, text):
        self.elements.append(Paragraph(f"• {text}", self.styles['KBullet']))
    
    def add_spacer(self, h=10):
        self.elements.append(Spacer(1, h))
    
    def add_line(self):
        self.elements.append(HRFlowable(width="100%", thickness=1, color=COLORS['light'],
                                        spaceBefore=8, spaceAfter=8))
    
    def add_page_break(self):
        self.elements.append(PageBreak())
    
    def add_chart(self, buf, caption=None, width=380, height=220):
        self.elements.append(Image(buf, width=width, height=height))
        if caption:
            self.elements.append(Paragraph(caption, self.styles['KCaption']))
        self.add_spacer(8)
    
    def add_highlight_box(self, text, color=None):
        """강조 박스"""
        if color is None:
            color = colors.HexColor('#EFF6FF')
        
        box_style = ParagraphStyle('BoxStyle', fontName=FONT, fontSize=10, leading=14,
                                   textColor=COLORS['dark'])
        content = [[Paragraph(text, box_style)]]
        box = Table(content, colWidths=[450])
        box.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), color),
            ('BOX', (0, 0), (-1, -1), 1, COLORS['primary']),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('RIGHTPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ]))
        self.elements.append(box)
        self.add_spacer(10)
    
    # ==========================================================================
    # 0. 목차 페이지
    # ==========================================================================
    def build_table_of_contents(self):
        """목차 페이지"""
        
        title_style = ParagraphStyle('TOCTitle', fontName=FONT_BOLD, fontSize=20, 
                                     alignment=TA_CENTER, textColor=COLORS['primary'], spaceAfter=25)
        self.elements.append(Paragraph("목 차", title_style))
        self.add_line()
        self.add_spacer(15)
        
        # 목차 항목 (목차가 2페이지이므로 실제 페이지 +1)
        toc_items = [
            ('1PAGE 요약', '핵심 결론과 전략 방향', 3),
            ('경영진 요약', '현황 진단 및 90일 로드맵', 4),
            ('2.1 PESTEL 분석', '거시환경 6대 영역 분석', 6),
            ('2.2 시나리오 분석', '미래 4대 시나리오', 8),
            ('2.3 경쟁환경 분석', 'Five Forces 및 경쟁사', 9),
            ('2.4 고객 분석', 'User/Payer/Beneficiary', 10),
            ('2.5 시장 분석', 'TAM/SAM/SOM 시장규모', 11),
            ('3.1 경영진단', '5대 영역 역량 평가', 12),
            ('3.2 VRIO 분석', '핵심 자원 경쟁우위', 13),
            ('3.3 SWOT 분석', '강점/약점/기회/위협', 14),
            ('3.4 TOWS 전략', '전략 옵션 및 우선순위', 15),
        ]
        
        # 목차 테이블
        toc_data = []
        for section, desc, page in toc_items:
            toc_data.append([section, desc, str(page)])
        
        toc_style = TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), FONT),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('FONTNAME', (0, 0), (0, -1), FONT_BOLD),
            ('TEXTCOLOR', (0, 0), (0, -1), COLORS['primary']),
            ('TEXTCOLOR', (1, 0), (1, -1), COLORS['gray']),
            ('TEXTCOLOR', (2, 0), (2, -1), COLORS['dark']),
            ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('LINEBELOW', (0, 0), (-1, -2), 0.5, colors.HexColor('#E5E7EB')),
        ])
        
        toc_table = Table(toc_data, colWidths=[150, 250, 50])
        toc_table.setStyle(toc_style)
        self.elements.append(toc_table)
        
        self.add_page_break()
    
    # ==========================================================================
    # 1. 1PAGE 요약 (한 페이지에 맞게 최적화)
    # ==========================================================================
    def build_one_page_summary(self):
        """1페이지 요약 - 핵심만 압축 (한 페이지 내 완결)"""
        
        # SWOT/TOWS에서 핵심 정보 추출
        swot = self.data.get('step_3_3_swot', {})
        tows = self.data.get('step_3_4_tows', {})
        mgmt = self.data.get('step_3_1_diagnosis', {})
        
        # 타이틀
        title_style = ParagraphStyle('OnepageTitle', fontName=FONT_BOLD, fontSize=16, 
                                     alignment=TA_CENTER, textColor=COLORS['primary'], spaceAfter=10)
        self.elements.append(Paragraph(f"{self.company_name} 분석 요약", title_style))
        self.add_line()
        
        # 핵심 결론 (간결하게)
        self.add_h3("핵심 결론")
        insights = swot.get('key_insights', [])
        if insights:
            for i, insight in enumerate(insights[:3], 1):
                short_insight = insight[:55] + '...' if len(insight) > 55 else insight
                self.add_small(f"<b>{i}.</b> {short_insight}")
        
        self.add_spacer(5)
        
        # 종합 진단 - 수평 막대 차트로 변경
        self.add_h3("종합 진단")
        scores = mgmt.get('scores_summary', {})
        if scores:
            # 수평 막대 점수 차트
            chart_buf = create_score_horizontal_bar(scores, width=420, height=130)
            self.add_chart(chart_buf, width=400, height=120)
        
        self.add_spacer(5)
        
        # TOP 3 전략 (테이블 작게)
        self.add_h3("핵심 전략 TOP 3")
        decision = tows.get('decision_summary', {})
        top_strategies = decision.get('top_3_strategies', [])
        if top_strategies:
            data = [['순위', '전략명', '유형', '핵심 근거']]
            for s in top_strategies[:3]:
                data.append([
                    str(s.get('rank', '')),
                    s.get('name', '')[:18],
                    s.get('type', ''),
                    s.get('rationale', '')[:28]
                ])
            table = styled_table(data, col_widths=[35, 120, 40, 200])
            self.elements.append(table)
        
        self.add_spacer(5)
        
        # 즉시 실행 과제 (2개만, 간결하게)
        self.add_h3("즉시 실행 과제")
        immediate = decision.get('immediate_actions', [])
        if immediate:
            action_data = [['과제', '담당', '기한']]
            for action in immediate[:2]:
                action_data.append([
                    action.get('action', '')[:35],
                    action.get('owner', ''),
                    action.get('deadline', '')
                ])
            action_table = Table(action_data, colWidths=[280, 60, 60])
            action_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), FONT),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD),
                ('BACKGROUND', (0, 0), (-1, 0), COLORS['warning']),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E5E7EB')),
                ('ROWHEIGHT', (0, 0), (-1, -1), 18),
            ]))
            self.elements.append(action_table)
        
        self.add_page_break()
    
    # ==========================================================================
    # 2. 경영진용 요약
    # ==========================================================================
    def build_executive_summary(self):
        """경영진용 요약"""
        self.add_h1("📈 경영진용 요약 (Executive Summary)")
        self.add_line()
        
        # 현황 진단
        self.add_h2("1. 현황 진단")
        
        pestel = self.data.get('step_2_1_pestel', {})
        summary = pestel.get('executive_summary', '')
        if summary:
            self.add_highlight_box(summary)
        
        # 외부환경 vs 내부역량
        self.add_h3("▶ 외부환경 (기회 vs 위협)")
        synthesis = pestel.get('synthesis', {})
        opportunities = synthesis.get('top_5_opportunities', [])[:2]
        threats = synthesis.get('top_5_threats', [])[:2]
        
        if opportunities:
            self.add_small("<font color='#10B981'><b>주요 기회:</b></font> " + 
                          " / ".join([o.get('factor', '')[:20] for o in opportunities]))
        if threats:
            self.add_small("<font color='#EF4444'><b>주요 위협:</b></font> " + 
                          " / ".join([t.get('factor', '')[:20] for t in threats]))
        
        self.add_spacer(8)
        
        self.add_h3("▶ 내부역량 (강점 vs 약점)")
        swot = self.data.get('step_3_3_swot', {})
        strengths = swot.get('strengths', [])[:2]
        weaknesses = swot.get('weaknesses', [])[:2]
        
        if strengths:
            self.add_small("<font color='#10B981'><b>핵심 강점:</b></font> " + 
                          " / ".join([s.get('description', '')[:25] for s in strengths]))
        if weaknesses:
            self.add_small("<font color='#EF4444'><b>핵심 약점:</b></font> " + 
                          " / ".join([w.get('description', '')[:25] for w in weaknesses]))
        
        self.add_spacer(15)
        
        # 전략 방향
        self.add_h2("2. 전략 방향")
        
        tows = self.data.get('step_3_4_tows', {})
        options = tows.get('strategy_options', {})
        
        for stype, label in [('SO', 'SO 전략 (공격)'), ('WO', 'WO 전략 (전환)'), 
                              ('ST', 'ST 전략 (방어)'), ('WT', 'WT 전략 (생존)')]:
            strategies = options.get(stype, [])
            if strategies:
                top = strategies[0]
                self.add_body(f"<b>{label}:</b> {top.get('name', '')} - {top.get('hypothesis', '')[:50]}...")
        
        self.add_spacer(15)
        
        # 실행 로드맵
        self.add_h2("3. 90일 실행 로드맵")
        
        sequencing = tows.get('strategy_sequencing', {})
        optimal = sequencing.get('optimal_sequence', {})
        
        if optimal:
            roadmap_data = [['단계', '기간', '핵심 전략', '목표']]
            for i, phase_key in enumerate(['phase_1', 'phase_2', 'phase_3'], 1):
                phase = optimal.get(phase_key, {})
                if phase:
                    roadmap_data.append([
                        f'Phase {i}',
                        phase.get('period', ''),
                        ', '.join(phase.get('strategies', [])),
                        phase.get('goals', '')[:35]
                    ])
            
            if len(roadmap_data) > 1:
                self.elements.append(styled_table(roadmap_data, col_widths=[55, 70, 100, 225]))
        
        self.add_spacer(15)
        
        # 리스크 관리
        self.add_h2("4. 핵심 리스크")
        
        risk_mgmt = tows.get('risk_management', {})
        pre_mortem = risk_mgmt.get('pre_mortem', [])
        
        if pre_mortem:
            risk_data = [['리스크', '발생확률', '예방조치']]
            for r in pre_mortem[:3]:
                risk_data.append([
                    r.get('failure_cause', '')[:30],
                    r.get('probability', ''),
                    r.get('preventive_action', '')[:35]
                ])
            self.elements.append(styled_table(risk_data, col_widths=[160, 60, 230], 
                                              header_color=COLORS['danger']))
        
        self.add_page_break()
    
    # ==========================================================================
    # 3. 단계별 상세 리포트
    # ==========================================================================
    def build_detailed_sections(self):
        """단계별 상세 리포트"""
        self.add_h1("📑 단계별 상세 분석")
        self.add_line()
        
        # 2.1 PESTEL
        self.build_pestel_detail()
        
        # 2.2 시나리오
        self.build_scenario_detail()
        
        # 2.3 경쟁환경
        self.build_competition_detail()
        
        # 2.4 고객분석
        self.build_customer_detail()
        
        # 2.5 시장분석
        self.build_market_detail()
        
        # 3.1 경영진단
        self.build_diagnosis_detail()
        
        # 3.2 VRIO
        self.build_vrio_detail()
        
        # 3.3 SWOT
        self.build_swot_detail()
        
        # 3.4 TOWS
        self.build_tows_detail()
    
    def build_pestel_detail(self):
        """2.1 PESTEL 상세"""
        self.add_h2("2.1 PESTEL 분석")
        
        pestel = self.data.get('step_2_1_pestel', {})
        pestel_data = pestel.get('pestel', {})
        
        # PESTEL 요약 히트맵 추가
        self._add_pestel_summary_chart(pestel_data)
        
        areas = [
            ('political', 'Political (정치)', 'P'),
            ('economic', 'Economic (경제)', 'E'),
            ('social', 'Social (사회)', 'S'),
            ('technological', 'Technological (기술)', 'T'),
            ('environmental', 'Environmental (환경)', 'E'),
            ('legal', 'Legal (법률)', 'L')
        ]
        
        for key, name, abbr in areas:
            area = pestel_data.get(key, {})
            if area:
                self.add_h3(f"▶ {name}")
                
                summary = area.get('summary', '')
                if summary:
                    self.add_small(f"<i>{summary}</i>")
                
                issues = area.get('issues', [])
                if issues:
                    issue_data = [['ID', '이슈', '영향', '긴급', '분류']]
                    for issue in issues[:4]:
                        issue_data.append([
                            issue.get('id', ''),
                            issue.get('name', '')[:20],
                            str(issue.get('impact_score', '')),
                            str(issue.get('urgency_score', '')),
                            issue.get('classification', '')
                        ])
                    self.elements.append(styled_table(issue_data, col_widths=[35, 180, 45, 45, 50], 
                                                      header_color=COLORS['pestel']))
                    self.add_spacer(8)
        
        # TOP 기회/위협
        synthesis = pestel.get('synthesis', {})
        
        self.add_h3("🌟 TOP 5 기회")
        opportunities = synthesis.get('top_5_opportunities', [])
        if opportunities:
            opp_data = [['순위', '영역', '요인', '실행방향']]
            for o in opportunities[:5]:
                opp_data.append([str(o.get('rank', '')), o.get('area', ''), 
                                o.get('factor', '')[:18], o.get('action', '')[:30]])
            self.elements.append(styled_table(opp_data, col_widths=[40, 60, 130, 220], 
                                              header_color=COLORS['accent']))
        
        self.add_spacer(8)
        
        self.add_h3("⚡ TOP 5 위협")
        threats = synthesis.get('top_5_threats', [])
        if threats:
            threat_data = [['순위', '영역', '요인', '대응방향']]
            for t in threats[:5]:
                threat_data.append([str(t.get('rank', '')), t.get('area', ''),
                                   t.get('factor', '')[:18], t.get('mitigation', '')[:30]])
            self.elements.append(styled_table(threat_data, col_widths=[40, 60, 130, 220],
                                              header_color=COLORS['danger']))
        
        self.add_page_break()
    
    def _add_pestel_summary_chart(self, pestel_data):
        """PESTEL 요약 차트 - 영역별 기회/위협 현황"""
        areas_info = [
            ('political', 'P', '정치'),
            ('economic', 'E', '경제'),
            ('social', 'S', '사회'),
            ('technological', 'T', '기술'),
            ('environmental', 'En', '환경'),
            ('legal', 'L', '법률')
        ]
        
        # 각 영역의 기회/위협 카운트
        summary_data = []
        for key, abbr, name in areas_info:
            area = pestel_data.get(key, {})
            issues = area.get('issues', [])
            opp_count = sum(1 for i in issues if i.get('classification') == '기회')
            threat_count = sum(1 for i in issues if i.get('classification') == '위협')
            # impact_score를 int로 변환
            impacts = []
            for i in issues:
                try:
                    impacts.append(int(i.get('impact_score', 0)))
                except:
                    impacts.append(0)
            avg_impact = sum(impacts) / len(impacts) if impacts else 0
            summary_data.append({
                'abbr': abbr, 'name': name,
                'opp': opp_count, 'threat': threat_count,
                'impact': avg_impact
            })
        
        # 테이블로 요약 표시
        table_data = [['영역', '기회', '위협', '영향도']]
        for s in summary_data:
            impact_bar = '●' * int(s['impact']) + '○' * (5 - int(s['impact']))
            table_data.append([
                f"{s['abbr']} ({s['name']})",
                str(s['opp']) if s['opp'] > 0 else '-',
                str(s['threat']) if s['threat'] > 0 else '-',
                impact_bar
            ])
        
        summary_table = Table(table_data, colWidths=[100, 50, 50, 80])
        summary_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), FONT),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#7C3AED')),  # 보라색
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('BACKGROUND', (1, 1), (1, -1), colors.HexColor('#ECFDF5')),  # 기회 - 연초록
            ('BACKGROUND', (2, 1), (2, -1), colors.HexColor('#FEF2F2')),  # 위협 - 연빨강
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E5E7EB')),
            ('ROWHEIGHT', (0, 0), (-1, -1), 20),
        ]))
        
        self.elements.append(summary_table)
        self.add_spacer(12)
    
    def build_scenario_detail(self):
        """2.2 시나리오 상세"""
        self.add_h2("2.2 시나리오 분석")
        
        scenario = self.data.get('step_2_2_scenario', {})
        
        # 시나리오 매트릭스
        scenarios_data = scenario.get('scenarios', {})
        scenarios_list = []
        for key in ['scenario_1', 'scenario_2', 'scenario_3', 'scenario_4']:
            s = scenarios_data.get(key, {})
            if s:
                scenarios_list.append({
                    'quadrant': s.get('quadrant', '++'),
                    'name': s.get('name', ''),
                    'probability': s.get('probability', '')
                })
        
        if scenarios_list:
            # 매트릭스와 확률 차트를 나란히 배치
            matrix_buf = create_scenario_matrix(scenarios_list)
            prob_buf = create_scenario_probability_chart(scenarios_list)
            
            # 2열 배치
            matrix_img = Image(matrix_buf, width=260, height=200)
            prob_img = Image(prob_buf, width=180, height=150)
            
            two_col = Table([[matrix_img, prob_img]], colWidths=[280, 200])
            two_col.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
            ]))
            self.elements.append(two_col)
            self.add_spacer(15)
        
        # 시나리오별 상세
        for key in ['scenario_1', 'scenario_2', 'scenario_3', 'scenario_4']:
            s = scenarios_data.get(key, {})
            if s:
                name = s.get('name', '')
                prob = s.get('probability', '')
                narrative = s.get('narrative', '')
                
                self.add_h3(f"● {name} ({prob})")
                self.add_small(narrative[:150] + '...' if len(narrative) > 150 else narrative)
                
                responses = s.get('strategic_response', [])
                if responses:
                    self.add_small(f"<b>대응전략:</b> {' / '.join(responses[:2])}")
                self.add_spacer(5)
        
        # 강건한 전략
        robust = scenario.get('robust_strategy', {})
        common = robust.get('common_strategies', [])
        if common:
            self.add_h3("🛡️ 강건한 전략 (모든 시나리오 공통)")
            for i, strategy in enumerate(common[:3], 1):
                self.add_bullet(strategy)
        
        self.add_page_break()
    
    def build_competition_detail(self):
        """2.3 경쟁환경 상세"""
        self.add_h2("2.3 경쟁환경 분석")
        
        competition = self.data.get('step_2_3_competition', {})
        five_forces = competition.get('five_forces', {})
        
        # Five Forces 차트
        if five_forces:
            chart_buf = create_five_forces_chart(five_forces)
            self.add_chart(chart_buf, width=280, height=280)
        
        # Five Forces 테이블
        overall = five_forces.get('overall', {})
        if overall:
            self.add_small(f"<b>산업 매력도:</b> {overall.get('industry_attractiveness', '')} (평균: {overall.get('average_score', '')}/5)")
        
        # 경쟁사 분석
        competitor_analysis = competition.get('competitor_analysis', {})
        competitors = competitor_analysis.get('business_competitors', [])
        
        if competitors:
            self.add_h3("주요 경쟁사")
            comp_data = [['경쟁사', '유형', '강점', '약점', '위협도']]
            for c in competitors[:4]:
                strengths = c.get('strengths', [''])[0][:15] if c.get('strengths') else ''
                weaknesses = c.get('weaknesses', [''])[0][:15] if c.get('weaknesses') else ''
                comp_data.append([
                    c.get('name', '')[:12], c.get('type', ''),
                    strengths, weaknesses, c.get('threat_level', '')
                ])
            self.elements.append(styled_table(comp_data, col_widths=[80, 60, 110, 110, 60]))
        
        self.add_page_break()
    
    def build_customer_detail(self):
        """2.4 고객분석 상세"""
        self.add_h2("2.4 고객 분석")
        
        customer = self.data.get('step_2_4_customer', {})
        ecosystem = customer.get('customer_ecosystem', {})
        
        for role, name in [('user', 'User (사용자)'), ('payer', 'Payer (지불자)'), ('beneficiary', 'Beneficiary (수혜자)')]:
            role_data = ecosystem.get(role, {})
            if role_data:
                self.add_h3(f"▶ {name}")
                profile = role_data.get('profile', '')
                self.add_body(f"<b>프로필:</b> {profile}")
                
                jtbd = role_data.get('jtbd', {})
                if jtbd:
                    functional = jtbd.get('functional', [])
                    if functional:
                        self.add_small(f"<b>JTBD:</b> {', '.join(functional[:3])}")
                self.add_spacer(5)
        
        # 세그먼트 우선순위
        priority = customer.get('segment_priority_matrix', {})
        if priority:
            self.add_h3("세그먼트 우선순위")
            priority_data = [['우선순위', '세그먼트', '선정이유', '접근전략']]
            for key in ['priority_1', 'priority_2']:
                p = priority.get(key, {})
                if p:
                    priority_data.append([
                        key.split('_')[1], p.get('segment', ''),
                        p.get('reason', '')[:25], p.get('approach_strategy', '')[:25]
                    ])
            if len(priority_data) > 1:
                self.elements.append(styled_table(priority_data, col_widths=[55, 100, 150, 145]))
        
        self.add_page_break()
    
    def build_market_detail(self):
        """2.5 시장분석 상세"""
        self.add_h2("2.5 시장 분석")
        
        market = self.data.get('step_2_5_market', {})
        sizing = market.get('market_sizing', {})
        
        # 시장 규모
        tam_data = sizing.get('tam', {})
        sam_data = sizing.get('sam', {})
        som_data = sizing.get('som', {})
        
        if tam_data:
            tri = tam_data.get('triangulation', {})
            tam = float(tri.get('confirmed_tam', 0))
            sam = float(sam_data.get('total', 0))
            som_y1 = float(som_data.get('year_1', {}).get('value', 0))
            
            if tam > 0:
                self.add_h3("시장 규모")
                # 동심원 차트 사용 - 정사각형 비율 유지
                chart_buf = create_concentric_market_chart(tam, sam, som_y1 if som_y1 > 0 else sam * 0.01)
                self.add_chart(chart_buf, width=320, height=320)  # 정사각형
        
        # 성장률
        trends = market.get('market_trends', {})
        growth = trends.get('growth_rates', {})
        if growth:
            self.add_h3("시장 성장률")
            growth_data = [['구분', '성장률', '기간']]
            for key, label in [('historical_cagr', '과거'), ('forecast_short', '단기'), ('forecast_mid', '중기')]:
                g = growth.get(key, {})
                if g:
                    growth_data.append([label, f"{g.get('value', '')}%", g.get('period', '')])
            if len(growth_data) > 1:
                self.elements.append(styled_table(growth_data, col_widths=[100, 100, 250]))
        
        self.add_page_break()
    
    def build_diagnosis_detail(self):
        """3.1 경영진단 상세"""
        self.add_h2("3.1 경영진단")
        
        diagnosis = self.data.get('step_3_1_diagnosis', {})
        
        # 요약
        summary = diagnosis.get('executive_summary', '')
        if summary:
            self.add_highlight_box(summary[:300] + '...' if len(summary) > 300 else summary)
        
        # 점수 차트 (레이더만)
        scores = diagnosis.get('scores_summary', {})
        if scores:
            chart_buf = create_diagnosis_radar_only(scores, width=280, height=280)
            self.add_chart(chart_buf, width=240, height=240)
            
            # 점수 테이블 (reportlab)
            score_data = [['영역', '점수', '상태', '핵심 평가']]
            for area, info in scores.items():
                score = float(info.get('score', 0))
                # 이모지 대신 텍스트 사용
                if score >= 4:
                    status = '양호'
                elif score >= 3:
                    status = '보통'
                else:
                    status = '취약'
                eval_text = info.get('evaluation', '')[:30]
                score_data.append([area, f'{score:.1f}', status, eval_text])
            self.elements.append(styled_table(score_data, col_widths=[70, 45, 40, 295]))
        
        self.add_page_break()
    
    def build_vrio_detail(self):
        """3.2 VRIO 상세"""
        self.add_h2("3.2 VRIO 분석")
        
        vrio = self.data.get('step_3_2_vrio', {})
        resources = vrio.get('resource_identification', {})
        resource_list = resources.get('resources', [])
        
        if resource_list:
            self.add_h3("핵심 자원")
            res_data = [['ID', '자원명', '유형', '신뢰도', '검증상태']]
            for r in resource_list[:5]:
                # 신뢰도 이모지를 텍스트로 변환
                reliability = r.get('final_reliability', '')
                if reliability in ['✅', 'verified']:
                    reliability_text = '높음'
                elif reliability in ['📊', 'partially_verified']:
                    reliability_text = '중간'
                elif reliability in ['⚠️', 'unverified']:
                    reliability_text = '낮음'
                else:
                    reliability_text = reliability[:4] if reliability else '-'
                
                # 검증상태 한글화
                status = r.get('verification_status', '')
                status_map = {'verified': '검증됨', 'partially_verified': '부분검증', 'unverified': '미검증'}
                status_text = status_map.get(status, status)
                
                res_data.append([
                    r.get('id', ''), r.get('name', '')[:18], r.get('type', ''),
                    reliability_text, status_text
                ])
            self.elements.append(styled_table(res_data, col_widths=[40, 140, 80, 50, 140],
                                              header_color=COLORS['vrio']))
        
        # VRIO 평가 시각화 차트 추가
        self._add_vrio_chart(resource_list)
        
        # 포트폴리오 요약
        portfolio = vrio.get('portfolio_summary', {})
        if portfolio:
            self.add_spacer(10)
            self.add_h3("경쟁 우위 포트폴리오")
            for key, label, color in [
                ('sustained_advantage', '지속적 경쟁우위', '#10B981'),
                ('temporary_advantage', '일시적 경쟁우위', '#3B82F6'),
                ('competitive_parity', '경쟁 균형', '#F59E0B')
            ]:
                items = portfolio.get(key, [])
                if items:
                    self.add_body(f"<font color='{color}'><b>{label}:</b></font> {', '.join(items)}")
        
        self.add_page_break()
    
    def _add_vrio_chart(self, resources):
        """VRIO 4요소 평가 차트"""
        if not resources:
            return
        
        # VRIO 평가 데이터 추출
        vrio_scores = {'V': 0, 'R': 0, 'I': 0, 'O': 0}
        count = 0
        
        for r in resources[:5]:
            vrio_eval = r.get('vrio_evaluation', {})
            if vrio_eval:
                count += 1
                vrio_scores['V'] += 1 if vrio_eval.get('valuable', {}).get('assessment') else 0
                vrio_scores['R'] += 1 if vrio_eval.get('rare', {}).get('assessment') else 0
                vrio_scores['I'] += 1 if vrio_eval.get('imitable', {}).get('assessment') else 0
                vrio_scores['O'] += 1 if vrio_eval.get('organized', {}).get('assessment') else 0
        
        if count == 0:
            return
        
        # 차트 데이터 (5개 자원 중 해당 요소를 충족하는 자원 수)
        chart_data = [
            ['요소', '설명', '충족 자원', '비율'],
            ['V', '가치(Valuable)', f'{vrio_scores["V"]}/{count}', f'{vrio_scores["V"]/count*100:.0f}%'],
            ['R', '희소성(Rare)', f'{vrio_scores["R"]}/{count}', f'{vrio_scores["R"]/count*100:.0f}%'],
            ['I', '모방난이도(Inimitable)', f'{vrio_scores["I"]}/{count}', f'{vrio_scores["I"]/count*100:.0f}%'],
            ['O', '조직화(Organized)', f'{vrio_scores["O"]}/{count}', f'{vrio_scores["O"]/count*100:.0f}%'],
        ]
        
        self.add_spacer(8)
        self.add_h3("VRIO 요소 충족 현황")
        
        vrio_table = Table(chart_data, colWidths=[40, 140, 80, 60])
        vrio_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), FONT),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD),
            ('BACKGROUND', (0, 0), (-1, 0), COLORS['vrio']),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E5E7EB')),
            ('ROWHEIGHT', (0, 0), (-1, -1), 22),
            # 비율에 따라 배경색
            ('BACKGROUND', (3, 1), (3, 1), colors.HexColor('#ECFDF5') if vrio_scores['V']/count >= 0.6 else colors.HexColor('#FEF2F2')),
            ('BACKGROUND', (3, 2), (3, 2), colors.HexColor('#ECFDF5') if vrio_scores['R']/count >= 0.6 else colors.HexColor('#FEF2F2')),
            ('BACKGROUND', (3, 3), (3, 3), colors.HexColor('#ECFDF5') if vrio_scores['I']/count >= 0.6 else colors.HexColor('#FEF2F2')),
            ('BACKGROUND', (3, 4), (3, 4), colors.HexColor('#ECFDF5') if vrio_scores['O']/count >= 0.6 else colors.HexColor('#FEF2F2')),
        ]))
        self.elements.append(vrio_table)
    
    def build_swot_detail(self):
        """3.3 SWOT 상세 - 2x2 매트릭스 레이아웃"""
        self.add_h2("3.3 SWOT 분석")
        
        swot = self.data.get('step_3_3_swot', {})
        
        # 각 사분면 데이터 준비
        def format_items(items, max_items=3):
            """아이템을 포맷팅"""
            formatted = []
            for item in items[:max_items]:
                desc = item.get('description', '')[:35]
                score = item.get('impact_score', '')
                formatted.append(f"• {desc} ({score})")
            return '\n'.join(formatted) if formatted else '-'
        
        s_items = format_items(swot.get('strengths', []))
        w_items = format_items(swot.get('weaknesses', []))
        o_items = format_items(swot.get('opportunities', []))
        t_items = format_items(swot.get('threats', []))
        
        # 2x2 매트릭스 테이블
        cell_style = ParagraphStyle('SWOTCell', fontName=FONT, fontSize=9, leading=12)
        
        matrix_data = [
            ['', '긍정적 요인', '부정적 요인'],
            ['내부\n환경', 
             Paragraph(f"<b><font color='#10B981'>강점 (S)</font></b><br/><br/>{s_items}", cell_style),
             Paragraph(f"<b><font color='#EF4444'>약점 (W)</font></b><br/><br/>{w_items}", cell_style)],
            ['외부\n환경',
             Paragraph(f"<b><font color='#3B82F6'>기회 (O)</font></b><br/><br/>{o_items}", cell_style),
             Paragraph(f"<b><font color='#F59E0B'>위협 (T)</font></b><br/><br/>{t_items}", cell_style)]
        ]
        
        matrix_table = Table(matrix_data, colWidths=[50, 200, 200], rowHeights=[25, 120, 120])
        matrix_table.setStyle(TableStyle([
            # 헤더
            ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD),
            ('FONTNAME', (0, 0), (0, -1), FONT_BOLD),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('FONTSIZE', (0, 0), (0, -1), 9),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F3F4F6')),
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F3F4F6')),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            # 셀 색상
            ('BACKGROUND', (1, 1), (1, 1), colors.HexColor('#ECFDF5')),  # S - 연초록
            ('BACKGROUND', (2, 1), (2, 1), colors.HexColor('#FEF2F2')),  # W - 연빨강
            ('BACKGROUND', (1, 2), (1, 2), colors.HexColor('#EFF6FF')),  # O - 연파랑
            ('BACKGROUND', (2, 2), (2, 2), colors.HexColor('#FFFBEB')),  # T - 연노랑
            # 테두리
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#D1D5DB')),
            ('BOX', (0, 0), (-1, -1), 2, COLORS['primary']),
            # 패딩
            ('LEFTPADDING', (1, 1), (-1, -1), 10),
            ('RIGHTPADDING', (1, 1), (-1, -1), 10),
            ('TOPPADDING', (1, 1), (-1, -1), 10),
            ('BOTTOMPADDING', (1, 1), (-1, -1), 10),
        ]))
        
        self.elements.append(matrix_table)
        self.add_spacer(15)
        
        # 핵심 인사이트
        insights = swot.get('key_insights', [])
        if insights:
            self.add_h3("핵심 인사이트")
            for i, insight in enumerate(insights[:3], 1):
                self.add_body(f"<b>{i}.</b> {insight[:80]}{'...' if len(insight) > 80 else ''}")
        
        # SWOT 요약 통계 추가
        self._add_swot_summary_stats(swot)
        
        self.add_page_break()
    
    def _add_swot_summary_stats(self, swot):
        """SWOT 요약 통계"""
        s_count = len(swot.get('strengths', []))
        w_count = len(swot.get('weaknesses', []))
        o_count = len(swot.get('opportunities', []))
        t_count = len(swot.get('threats', []))
        
        # 평균 영향도 계산
        def avg_impact(items):
            scores = [item.get('impact_score', 0) for item in items]
            return sum(scores) / len(scores) if scores else 0
        
        s_avg = avg_impact(swot.get('strengths', []))
        w_avg = avg_impact(swot.get('weaknesses', []))
        o_avg = avg_impact(swot.get('opportunities', []))
        t_avg = avg_impact(swot.get('threats', []))
        
        self.add_spacer(10)
        self.add_h3("SWOT 요약 통계")
        
        stats_data = [
            ['구분', '항목 수', '평균 영향도', '분석 결과'],
            ['강점 (S)', str(s_count), f'{s_avg:.1f}/5', '핵심 경쟁력' if s_avg >= 4 else '보통'],
            ['약점 (W)', str(w_count), f'{w_avg:.1f}/5', '심각' if w_avg >= 4 else '관리 필요'],
            ['기회 (O)', str(o_count), f'{o_avg:.1f}/5', '적극 활용' if o_avg >= 4 else '선별 활용'],
            ['위협 (T)', str(t_count), f'{t_avg:.1f}/5', '즉시 대응' if t_avg >= 4 else '모니터링'],
        ]
        
        stats_table = Table(stats_data, colWidths=[80, 70, 80, 100])
        stats_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), FONT),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD),
            ('BACKGROUND', (0, 0), (-1, 0), COLORS['swot']),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('BACKGROUND', (0, 1), (0, 1), colors.HexColor('#ECFDF5')),  # S
            ('BACKGROUND', (0, 2), (0, 2), colors.HexColor('#FEF2F2')),  # W
            ('BACKGROUND', (0, 3), (0, 3), colors.HexColor('#EFF6FF')),  # O
            ('BACKGROUND', (0, 4), (0, 4), colors.HexColor('#FFFBEB')),  # T
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E5E7EB')),
            ('ROWHEIGHT', (0, 0), (-1, -1), 22),
        ]))
        self.elements.append(stats_table)
    
    def build_tows_detail(self):
        """3.4 TOWS 상세"""
        self.add_h2("3.4 TOWS 전략")
        
        tows = self.data.get('step_3_4_tows', {})
        options = tows.get('strategy_options', {})
        
        # 전략 유형별 테이블로 정리
        strategy_data = []
        
        for stype, label, desc in [
            ('SO', 'SO 전략', '강점으로 기회 포착'),
            ('WO', 'WO 전략', '약점 보완하며 기회 활용'),
            ('ST', 'ST 전략', '강점으로 위협 방어'),
            ('WT', 'WT 전략', '약점/위협 최소화')
        ]:
            strategies = options.get(stype, [])
            for s in strategies[:2]:
                name = s.get('name', '')
                hypothesis = s.get('hypothesis', '')[:50]
                eval_data = s.get('evaluation', {})
                total_score = eval_data.get('total_score', 0)
                
                # 우선순위 표시
                if total_score >= 22:
                    priority = '★★★'
                elif total_score >= 20:
                    priority = '★★'
                else:
                    priority = '★'
                
                strategy_data.append({
                    'type': stype,
                    'name': name,
                    'hypothesis': hypothesis,
                    'score': total_score,
                    'priority': priority
                })
        
        # 전략 테이블
        if strategy_data:
            table_data = [['유형', '전략명', '핵심 가설', '점수', '우선순위']]
            for s in strategy_data:
                table_data.append([
                    s['type'], s['name'][:18], s['hypothesis'] + '...', 
                    str(s['score']), s['priority']
                ])
            self.elements.append(styled_table(table_data, col_widths=[40, 120, 180, 40, 70]))
        
        self.add_spacer(15)
        
        # 최종 전략 우선순위
        decision = tows.get('decision_summary', {})
        top = decision.get('top_3_strategies', [])
        if top:
            self.add_h3("최종 전략 우선순위")
            top_data = [['순위', '전략', '유형', '선정 근거']]
            for s in top[:3]:
                top_data.append([
                    str(s.get('rank', '')), s.get('name', '')[:20],
                    s.get('type', ''), s.get('rationale', '')[:35]
                ])
            self.elements.append(styled_table(top_data, col_widths=[40, 140, 50, 220], 
                                              header_color=COLORS['secondary']))
        
        # 즉시 실행 과제
        immediate = decision.get('immediate_actions', [])
        if immediate:
            self.add_spacer(15)
            self.add_h3("즉시 실행 과제")
            for action in immediate[:3]:
                task = action.get('action', '')
                owner = action.get('owner', '')
                deadline = action.get('deadline', '')
                self.add_body(f"• {task}")
                self.add_small(f"   담당: {owner} / 기한: {deadline}")
        
        # 전략 로드맵 타임라인
        if top:
            self.add_spacer(15)
            self.add_h3("전략 실행 로드맵")
            roadmap_buf = create_strategy_roadmap(top)
            self.add_chart(roadmap_buf, width=440, height=160)
    
    # ==========================================================================
    # 빌드
    # ==========================================================================
    def build(self):
        """전체 빌드"""
        # 표지 후 빈 콘텐츠로 페이지 넘김 처리
        # (표지는 onFirstPage에서 그려지므로 첫 element 전에 PageBreak 불필요)
        
        self.build_table_of_contents()  # 목차 추가
        self.build_one_page_summary()
        self.build_executive_summary()
        self.build_detailed_sections()
        return self.elements


# ==============================================================================
# 메인 함수
# ==============================================================================
def generate_analysis_report(data, output_path, company_name=None):
    """분석 리포트 PDF 생성"""
    
    if company_name is None:
        pestel = data.get('step_2_1_pestel', {})
        meta = pestel.get('analysis_meta', {})
        company_name = meta.get('company', '기업명')
    
    report_date = datetime.now().strftime('%Y년 %m월 %d일')
    template = ReportTemplate(company_name, report_date)
    
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        rightMargin=15*mm, leftMargin=15*mm,
        topMargin=25*mm, bottomMargin=20*mm
    )
    
    builder = AnalysisReportBuilder(data, company_name)
    
    # 표지 후 콘텐츠 시작을 위해 빈 요소 + PageBreak 추가
    from reportlab.platypus import PageBreak, Spacer
    cover_elements = [Spacer(1, 1), PageBreak()]  # 표지 페이지 채우기
    
    content_elements = builder.build()
    all_elements = cover_elements + content_elements
    
    doc.build(all_elements, onFirstPage=template.cover_page, onLaterPages=template.header_footer)
    
    return output_path


# ==============================================================================
# 테스트
# ==============================================================================
if __name__ == '__main__':
    # 샘플 데이터
    sample = {
        'step_2_1_pestel': {
            'analysis_meta': {'company': 'G임팩트'},
            'executive_summary': '2026년 G임팩트는 정부의 AI·딥테크 육성 정책과 지역 소멸 대응 기조에 힘입어 강력한 성장 기회를 맞이했습니다. SLM 기술을 통한 내부 리소스 문제 해결이 가능하나, 벤처 투자 침체와 AI 규제 강화, 그리고 심각한 내부 번아웃이 주요 위협입니다.',
            'pestel': {
                'political': {
                    'summary': '정부의 AI 육성 및 지역 균형 발전 정책은 호재이나, 높은 B2G 의존도는 리스크.',
                    'issues': [
                        {'id': 'P1', 'name': 'AI·딥테크 스타트업 육성 정책', 'impact_score': 5, 'urgency_score': 4, 'classification': '기회'},
                        {'id': 'P2', 'name': '비수도권 정책금융 확대', 'impact_score': 4, 'urgency_score': 3, 'classification': '기회'},
                    ]
                }
            },
            'synthesis': {
                'top_5_opportunities': [
                    {'rank': 1, 'area': 'Political', 'factor': '정부의 AI·딥테크 육성 정책', 'impact_score': 5, 'action': '지역 특화 AI 사업 수주'},
                    {'rank': 2, 'area': 'Tech', 'factor': 'SLM 기술 효율화', 'impact_score': 5, 'action': '저비용 개발'},
                ],
                'top_5_threats': [
                    {'rank': 1, 'area': 'Legal', 'factor': 'AI 규제 강화', 'urgency_score': 5, 'mitigation': '데이터 비식별화'},
                    {'rank': 2, 'area': 'Social', 'factor': '내부 번아웃', 'urgency_score': 5, 'mitigation': '강제 휴식'},
                ]
            }
        },
        'step_2_2_scenario': {
            'scenarios': {
                'scenario_1': {'quadrant': '++', 'name': '황금기', 'probability': '20%', 'narrative': '정부의 전폭적인 지원과 지역 생태계 활황', 'strategic_response': ['공격적 Scale-up']},
                'scenario_2': {'quadrant': '-+', 'name': '지역의 봄', 'probability': '30%', 'narrative': '정부 지원 감소, 민간 자생', 'strategic_response': ['Pivot to Private']},
                'scenario_3': {'quadrant': '--', 'name': '빙하기', 'probability': '15%', 'narrative': '최악의 상황', 'strategic_response': ['비상 경영']},
                'scenario_4': {'quadrant': '+-', 'name': '수도권 독주', 'probability': '35%', 'narrative': '수도권 집중', 'strategic_response': ['Hybrid Operation']},
            },
            'robust_strategy': {'common_strategies': ['B2G 의존도 50% 이하', 'SLM 도입', '투자 네트워크 구축']}
        },
        'step_2_3_competition': {
            'five_forces': {
                'new_entrants': {'score': 3}, 'rivalry': {'score': 4.5}, 'substitutes': {'score': 2.5},
                'supplier_power': {'score': 4}, 'buyer_power': {'score': 4},
                'overall': {'average_score': 3.6, 'industry_attractiveness': '중간'}
            },
            'competitor_analysis': {
                'business_competitors': [
                    {'name': 'SKT AI Lab', 'type': '잠재경쟁', 'strengths': ['기술 인프라'], 'weaknesses': ['지역 이해도 낮음'], 'threat_level': '중간'},
                    {'name': '이드로', 'type': '직접경쟁', 'strengths': ['전남 기반'], 'weaknesses': ['AI 역량 부재'], 'threat_level': '높음'},
                ]
            }
        },
        'step_2_4_customer': {
            'customer_ecosystem': {
                'user': {'profile': '지역 기반 초기 혁신가', 'jtbd': {'functional': ['투자 유치', 'IR 자료 작성']}},
                'payer': {'profile': '예산 효율성 필요한 조직 리더', 'jtbd': {'functional': ['성과 리포팅']}},
                'beneficiary': {'profile': '지역 사회', 'jtbd': {}}
            },
            'segment_priority_matrix': {
                'priority_1': {'segment': '지자체/공공기관', 'reason': '검증된 예산', 'approach_strategy': '성과 관리 툴'},
                'priority_2': {'segment': '대기업 ESG팀', 'reason': '높은 객단가', 'approach_strategy': '협력사 관리'}
            }
        },
        'step_2_5_market': {
            'market_sizing': {
                'tam': {'triangulation': {'confirmed_tam': 120000}},
                'sam': {'total': 9289},
                'som': {'year_1': {'value': 20}}
            },
            'market_trends': {
                'growth_rates': {
                    'historical_cagr': {'value': 18, 'period': '2021-2026'},
                    'forecast_short': {'value': 5.2, 'period': '1년'},
                    'forecast_mid': {'value': 15, 'period': '3년'}
                }
            }
        },
        'step_3_1_diagnosis': {
            'executive_summary': 'G임팩트는 명확한 소셜 미션을 보유하나, 경영일반(2.6점)과 인사노무(3.0점) 영역에서 시스템 부재가 성장을 가로막고 있습니다.',
            'scores_summary': {
                '사회적가치': {'score': '3.1', 'evaluation': '미션 명확, 협력 네트워크 부족'},
                '경영일반': {'score': '2.6', 'evaluation': '가장 취약. 피드백 시스템 부재'},
                '영업마케팅': {'score': '3.3', 'evaluation': '분석 우수, 실행력 보완 필요'},
                '재무': {'score': '4.0', 'evaluation': '관리 양호, 정부 의존도 높음'},
                '인사조직': {'score': '3.0', 'evaluation': '피로도 누적, 리더십 신뢰 부족'}
            }
        },
        'step_3_2_vrio': {
            'resource_identification': {
                'resources': [
                    {'id': 'R1', 'name': 'MYSC JV 파트너십', 'type': 'relational', 'final_reliability': '✅', 'verification_status': 'verified'},
                    {'id': 'R2', 'name': '광주·전남 로컬 네트워크', 'type': 'relational', 'final_reliability': '📊', 'verification_status': 'partial'},
                    {'id': 'R3', 'name': 'AI 엔진 4종', 'type': 'technological', 'final_reliability': '⚠️', 'verification_status': 'unverified'},
                ]
            },
            'portfolio_summary': {
                'sustained_advantage': [],
                'temporary_advantage': ['광주·전남 로컬 네트워크'],
                'competitive_parity': ['AI 엔진 4종']
            }
        },
        'step_3_3_swot': {
            'key_insights': [
                'MYSC 파트너십과 정부 예산을 결합하여 단기 런웨이 확보가 급선무',
                '조직 운영 시스템 개선 없이는 어떤 전략도 지속 불가능',
                '대기업 ESG 시장 진입으로 B2G 의존도와 투자 침체를 동시 해결'
            ],
            'strengths': [
                {'id': 'S1', 'description': 'MYSC JV 파트너십', 'impact_score': 5, 'priority': '⭐⭐⭐'},
                {'id': 'S2', 'description': '광주·전남 로컬 네트워크', 'impact_score': 4, 'priority': '⭐⭐⭐'},
            ],
            'weaknesses': [
                {'id': 'W1', 'description': '조직 운영 시스템 부재 & 리더십 번아웃', 'impact_score': 5, 'priority': '⭐⭐⭐'},
                {'id': 'W2', 'description': '높은 B2G 의존도 (80%)', 'impact_score': 5, 'priority': '⭐⭐⭐'},
            ],
            'opportunities': [
                {'id': 'O1', 'description': '정부 지역소멸 대응 예산 확대', 'impact_score': 5, 'priority': '⭐⭐⭐'},
                {'id': 'O2', 'description': '대기업 ESG 실사 의무화', 'impact_score': 5, 'priority': '⭐⭐⭐'},
            ],
            'threats': [
                {'id': 'T1', 'description': '벤처투자 시장 침체', 'impact_score': 5, 'priority': '⭐⭐⭐'},
                {'id': 'T2', 'description': '수도권 대형 AC의 지역 진출', 'impact_score': 4, 'priority': '⭐⭐⭐'},
            ]
        },
        'step_3_4_tows': {
            'strategy_options': {
                'SO': [{'name': '로컬 임팩트 메가 프로젝트', 'hypothesis': 'MYSC와 컨소시엄 구성으로 대형 지자체 사업 수주', 'evaluation': {'total_score': 24, 'priority': '⭐⭐⭐'}}],
                'WO': [{'name': '공공 자금 기반 조직 시스템화', 'hypothesis': '정부 사업비로 운영 시스템 구축', 'evaluation': {'total_score': 22, 'priority': '⭐⭐⭐'}}],
                'ST': [{'name': '로컬 데이터 장벽 구축', 'hypothesis': '지역 데이터 축적으로 진입장벽 형성', 'evaluation': {'total_score': 21, 'priority': '⭐⭐'}}],
                'WT': [{'name': '생존형 피봇 & 리더십 케어', 'hypothesis': '불필요한 R&D 중단, 휴식', 'evaluation': {'total_score': 22, 'priority': '⭐⭐'}}],
            },
            'strategy_sequencing': {
                'optimal_sequence': {
                    'phase_1': {'period': '0-6개월', 'strategies': ['WO-1', 'WT-1'], 'goals': '조직 안정화 및 번아웃 해소'},
                    'phase_2': {'period': '6-12개월', 'strategies': ['SO-1'], 'goals': '대형 공공 사업 수주'},
                    'phase_3': {'period': '1-2년', 'strategies': ['SO-2', 'ST-1'], 'goals': 'B2B SaaS 런칭'}
                }
            },
            'decision_summary': {
                'top_3_strategies': [
                    {'rank': 1, 'name': '공공 자금 기반 조직 시스템화', 'type': 'WO', 'rationale': '조직 붕괴 리스크 해소 및 실행 기반 마련'},
                    {'rank': 2, 'name': '로컬 임팩트 메가 프로젝트', 'type': 'SO', 'rationale': '매출 증대 및 시장 지배력 확대'},
                    {'rank': 3, 'name': 'AI ESG 공급망 솔루션', 'type': 'SO', 'rationale': '수익 다변화 및 스케일업'},
                ],
                'immediate_actions': [
                    {'action': 'CEO 주말 근무 중단 및 업무 위임 리스트 작성', 'owner': 'CEO', 'deadline': '1주 내'},
                    {'action': 'MYSC 담당자와 공동 사업 기획 미팅', 'owner': 'CSO', 'deadline': '2주 내'},
                ]
            },
            'risk_management': {
                'pre_mortem': [
                    {'failure_cause': 'MYSC 협력이 MOU 수준에 그침', 'probability': '중간', 'preventive_action': '협력 전담자 지정 및 월 1회 정기 교류'},
                    {'failure_cause': '리더십 번아웃으로 의사결정 지연', 'probability': '높음', 'preventive_action': '강제 휴가 및 R&R 재분배'},
                ]
            }
        }
    }
    

# 실제 데이터로 테스트
if __name__ == "__main__":
    import json
    
    # 실제 샘플 데이터 로드
    with open('real_sample_data.json', 'r', encoding='utf-8') as f:
        real_data = json.load(f)
    
    output_path = '/home/claude/G임팩트_분석리포트_실제데이터.pdf'
    generate_analysis_report(real_data, output_path, 'G임팩트')
    print(f"PDF 생성 완료: {output_path}")
