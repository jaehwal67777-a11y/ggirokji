# -*- coding: utf-8 -*-
"""
주간활동서비스(성인) 월별 활동계획 및 기록지 자동 작성기  -  노원점부터 시작
(방과후 엔진 make_records.py 는 건드리지 않고, 공통 도구만 빌려 씁니다)

필요한 파일 3개:  ① 결제내역 엑셀  ② 시간표 엑셀  ③ 주간활동서비스 빈 양식 엑셀
결과:  주간활동_N월_계획서_기록지.xlsx (아동마다 '계획서' '기록지' 두 시트) + 점검표.xlsx

지점마다 다른 점은 지점설정.txt 에서 바꿔요 (줄 맨 앞 # 은 설명)
"""
import os
import re
import sys
import math
import datetime as dt
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill


# ---------------------------------------------------------------------
# 방과후 엔진의 공통 도구(엑셀 열기, 휴일 계산 등)를 빌려 옵니다
# ---------------------------------------------------------------------
def _기본엔진_찾기():
    후보 = []
    try:
        후보.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'make_records.py'))
    except Exception:
        pass
    후보 += ['/engine/make_records.py', os.path.join(os.getcwd(), 'make_records.py')]
    for p in 후보:
        if os.path.exists(p):
            return p
    raise SystemExit('공통 도구 파일(make_records.py)을 찾지 못했어요.')


_기본경로 = _기본엔진_찾기()
B = {'__name__': 'base_engine', '__file__': os.path.join(os.getcwd(), 'make_records.py')}
exec(compile(open(_기본경로, encoding='utf-8').read(), _기본경로, 'exec'), B)
엑셀열기 = B['엑셀열기']
줄수 = B['줄수']
글자폭 = B['글자폭']

# =====================================================================
# [설정]  지점설정.txt 로 바꿀 수 있어요
# =====================================================================
연도 = None                    # None = 자동
칸_구분 = ' / '                   # 칸과 칸 사이 (양천: 줄바꿈 '\n')
특이_줄 = False                   # True = 개별 특이사항을 ' -내용' 줄바꿈 모양으로 (양천)
슬래시_고르기 = True              # False = 시간표에 'A/B'처럼 적힌 칸을 하나만 고르지 않고 그대로 써요 (양천 최중증: 특수체육/감각통합)
소제목_줄 = False                  # True = 대제목 줄 아래에 ' -소제목' 줄을 따로 (양천 최중증)
인정결제_문구 = '인정결제 {n}시간'       # 양천: '인정결제({n}시간)'
괄호붙이기 = False                # True = '활동 (소제목)'의 괄호 앞 공백을 없애요 → 활동(소제목) (양천)
글자_바꾸기 = {}                   # 예: {'스마트교육센터(AI체험 및 VR)': '스마트교육센터(AI체험 및 VR)(협)'}
고정_칸 = []                      # 예: [(9.5, 10), (10, 11.5), (11.5, 13), (13, 14), (14, 15)] 칸 경계가 1시간이 아닐 때 (양천 최중증)
시각_표기 = '시'                   # '시' = 9시~10시 / '콜론' = 09:00~10:00
송영_시간에_포함 = False          # True = 송영 칸을 활동 시작 바로 1시간 전에 넣어요 (송영 09~10시, 활동 10시부터, 양천점)
상세로_대체 = False              # True = 시간표 아래칸 세부내용이 있으면 위칸 이름 대신 세부내용만 써요 (양천점)
빈_요구사항_줄_빼기 = False       # True = 요구사항이 비어 있으면 요구사항 빈 줄을 만들지 않아요 (양천점)
시간표_새모양 = False            # True = 시간표 첫 칸이 '09:00 - 10:00' / '9/1 - 4 (1주)' 모양 (양천점)
사업유형_글자 = ''               # 결제내역에 사업유형이 섞여 있을 때, 이 글자가 들어 있는 사업유형만 써요 (예: '통합돌봄')
출력_폴더 = '주간활동기록지'
지점이름 = ''
수급시간_기본형 = 132           # 기본형 수급시간 (결제내역 등급에 '기본'이 들어 있으면)
수급시간_확장형 = 176           # 확장형 수급시간 ('확장'이 들어 있으면)
외부_최소비율 = 0.30            # (자체 외부 + 협력) ÷ 수급시간 이 이 비율 이상이어야 해요
하루_마지막_시각 = 18           # 하루 활동이 이 시각을 넘지 않게 맞춰요
점심_문구 = ''                  # 비우면 시간표 글 그대로 / 적으면 12~13시 칸에 이 글을 써요 (예: '점심식사')
송영_문구 = '송영서비스'
외부활동_글자 = ['지역사회', '산책', '외부', '나들이', '견학', '현장학습', '맛집', '놀이터', '탐방', '마트']
협력_글자 = ['(협)', '협력']
상위이름_생략 = ['특별활동']     # 시간표 위칸이 이 이름이면, 아래칸 글만 써요
보조_활동 = ['자율활동(젠가)', '자율활동(퍼즐)', '개인학습', '컴퓨터활용', '자율활동(색칠공부)', '자율활동(미로찾기)', '자율활동(숨은그림찾기)']
보조_외부활동 = '놀이터'         # 마지막 시간이 시간표 밖일 때 쓰는 센터 밖 활동
외부_채울_활동 = '산책'          # 30%가 모자랄 때 바꿔 넣는 센터 밖 활동
바꿔도_되는_글자 = ['자율활동', '개인학습', '컴퓨터활용', '영화감상', '노래방', '명상']   # 30%가 모자랄 때 산책으로 바꿔도 되는 활동
묶음_지정 = {}                  # 시간표의 두 번째 묶음(아래쪽 표)을 쓸 아동. 예: {'김하나': 'B'}  (기본은 모두 첫 번째 표)
송영_있는날_두번째표 = True      # 아침 송영이 있는 날은 두 번째 표(10시 시작)를 써요
축소사유_지정 = {}              # 예: {'김하나': '탁구장'}
축소사유_돌려쓰기 = []          # 여러 사유를 아동 순서대로 돌아가며 넣어요. 예: ['- 치료 일정', '- 학원 일정']
계획시간_축소사유 = ''          # 모든 아동에게 같은 말을 쓸 때
_돌림_순번 = {}


def _돌려쓰기_고르기(아동):
    if 축소사유_돌려쓰기:
        if 아동 not in _돌림_순번:
            _돌림_순번[아동] = len(_돌림_순번)
        return 축소사유_돌려쓰기[_돌림_순번[아동] % len(축소사유_돌려쓰기)]
    return 계획시간_축소사유

담당_지정 = {}                  # 예: {'김하나': '이숙'}
대체인력_문구 = '{담당} 선생님 연차로 제공인력 변경({대체})'
늦은등원_사유 = '늦잠'            # 평소보다 늦게 와서 시간이 줄었을 때 특이사항에 쓸 사유
추가이용_사유 = ['보호자 요청', '이용자 요청']   # 평소보다 오래 이용한 날 사유 (여러 개면 날짜마다 돌아가며 써요)
불참_사유 = ['이용자 컨디션 난조', '활동보조 일정 조율', '학부모 일정 조율']   # 결제가 없는 평일(센터는 열었는데 안 온 날) 사유
불참_문구 = '불참 (사유:{사유})'   # 결제가 없는 평일 특이사항 모양. 예: '결석({사유})'
불참_지정 = {}                  # 날짜별 결석 사유를 직접 정해요. 예: {'홍길동': {'9/3': '병원', '9/10': '부모님 요청'}}  (안 적은 날은 불참_사유대로)
송영_대상자 = []                # 비우면 송영 결제를 모두 쓰고, 이름을 적으면 그 사람들만 송영을 써요. 예: ['염지훈', '홍민경']
월간_활동 = []                  # 한 달에 한두 번 하는 활동. 아래 설명 참고
이른하원_사유 = '병원'            # 평소보다 일찍 가서 시간이 줄었을 때 쓸 사유
시간변동_표기 = True             # 평소와 다른 이용시간인 날 '추가이용/단축'을 특이사항에 써요
소급_문구 = {'대상자 과실(카드 미소지 등)': '카드 미소지로 소급결제',
            '제공인력 과실(단말기 조작미숙 등)': '제공인력 결제미숙으로 소급결제',
            '서버 접속 실패(0104 오류 등)': '서버 접속 실패로 소급결제'}
시트이름_구분 = ' '
추가_공휴일 = []
요구사항_머리말 = '이용자 요구사항) '
기록일_방식 = '인정결제'         # '인정결제' = 인정결제된 날 / '말일' = 그 달 마지막 근무일
인정결제_요구사항에 = False        # True = 인정결제 시간을 요구사항 줄 끝에 적어요 / False(기본) = 인정결제한 날짜 줄의 개별 특이사항에 적어요

설정_이름들 = ['인정결제_요구사항에', '시간표_새모양', '송영_시간에_포함', '고정_칸', '시각_표기', '칸_구분', '특이_줄', '소제목_줄', '슬래시_고르기', '괄호붙이기', '인정결제_문구', '글자_바꾸기', '상세로_대체', '빈_요구사항_줄_빼기', '사업유형_글자', '연도', '출력_폴더', '지점이름', '수급시간_기본형', '수급시간_확장형', '외부_최소비율', '하루_마지막_시각', '점심_문구', '송영_문구',
            '외부활동_글자', '협력_글자', '상위이름_생략', '보조_활동', '보조_외부활동', '외부_채울_활동', '바꿔도_되는_글자', '묶음_지정',
            '송영_있는날_두번째표', '축소사유_지정', '축소사유_돌려쓰기', '계획시간_축소사유', '담당_지정', '대체인력_문구', '시간변동_표기', '늦은등원_사유', '이른하원_사유', '추가이용_사유', '불참_사유', '불참_문구', '불참_지정', '송영_대상자', '월간_활동', '소급_문구',
            '시트이름_구분', '추가_공휴일', '요구사항_머리말', '기록일_방식']
설정_설명 = {
    '수급시간_기본형': '기본형 수급시간(시간). 결제내역 등급에 "기본"이 들어 있으면 이 값',
    '수급시간_확장형': '확장형 수급시간(시간). 등급에 "확장"이 들어 있으면 이 값',
    '외부_최소비율': '(자체 외부 + 협력기관) 시간이 수급시간의 이 비율 이상이어야 해요 (0.30 = 30%)',
    '하루_마지막_시각': '하루 활동이 이 시각을 넘지 않게 맞춰요 (18 = 오후 6시)',
    '점심_문구': "비우면 시간표의 12~13시 글을 그대로 써요. 적으면 그 글로 바꿔요. 예: '점심식사'",
    '외부활동_글자': "활동 이름에 이 글자가 들어 있으면 센터 밖(외부) 활동으로 세요",
    '협력_글자': "활동 이름에 이 글자가 있으면 협력기관 시간으로 세요",
    '묶음_지정': "시간표 아래쪽(두 번째) 표를 쓸 아동. 예: {'김하나': 'B'}",
    '축소사유_돌려쓰기': "여러 사유를 아동 순서대로 돌아가며 넣어요. 예: ['- 치료 일정', '- 학원 일정', '- 활동 보조인']",
    '축소사유_지정': "아동별 '계획시간 축소 사유'. 예: {'김하나': '탁구장'}",
    '담당_지정': "담당 선생님을 직접 정해요. 예: {'김하나': '이숙'}",
    '보조_활동': "시간표가 비었거나 '자율활동'만 적힌 시간에 돌아가며 넣는 활동. 여러 개 적을 수 있어요. 예: ['자율활동(퍼즐)', '자율활동(젠가)', '개인학습', '컴퓨터활용']",
    '늦은등원_사유': '평소보다 늦게 와서 시간이 줄어든 날 특이사항에 쓸 사유',
    '추가이용_사유': "평소보다 오래 이용한 날 특이사항 사유. 여러 개면 돌아가며 써요. 예: ['보호자 요청', '이용자 요청']",
    '불참_문구': "결제가 없는 평일 특이사항 모양. {사유}가 사유로 바뀌어요. 예: '결석({사유})'",
    '불참_지정': "날짜별 결석 사유를 직접 정해요. 예: {'홍길동': {'9/3': '병원', '9/10': '부모님 요청'}}",
    '송영_대상자': "이름을 적으면 그 사람들만 송영을 써요. 비우면 모두. 예: ['염지훈', '홍민경']",
    '불참_사유': "센터는 열었는데 결제가 없는 평일의 사유. 여러 개면 돌아가며 써요. 예: ['이용자 컨디션 난조', '학부모 일정 조율']",
    '월간_활동': "한 달에 한두 번 하는 활동. 예: [{'이름': '클라이밍', '요일': '목', '시작': 10, '끝': 11, '주차': 2, '이용자': ['홍길동', '김하나']}]  (이용자를 비우고 '인원': 5 로 적으면 5명을 알아서 골라요). 특정 날짜 하루만이면 '요일','주차' 대신 '날짜': '10/22' 로 적고, 이용자를 비우면 모두예요. 예: [{'이름': '밤따기', '날짜': '10/22', '시작': 9, '끝': 15}]",
    '이른하원_사유': '평소보다 일찍 가서 시간이 줄어든 날 특이사항에 쓸 사유',
    '대체인력_문구': '담당이 아닌 선생님이 결제한 날 특이사항 글. {담당} {대체} 가 이름으로 바뀌어요',
}


def 설정_불러오기():
    import ast
    g = globals()
    바뀜, 문제 = [], []
    폴더들 = list(dict.fromkeys([os.getcwd(), os.path.dirname(os.path.abspath(__file__)) if '__file__' in g else os.getcwd()]))
    for 폴더 in 폴더들:
        경로 = os.path.join(폴더, '지점설정.txt')
        if not os.path.exists(경로):
            continue
        try:
            글 = open(경로, encoding='utf-8-sig').read()
        except Exception:
            글 = open(경로, encoding='cp949', errors='replace').read()
        for 번호, 줄 in enumerate(글.splitlines(), 1):
            줄 = 줄.strip()
            if not 줄 or 줄.startswith('#') or '=' not in 줄:
                continue
            이름, _, 값 = 줄.partition('=')
            이름 = 이름.strip()
            if 이름 not in 설정_이름들:
                문제.append(f'지점설정.txt {번호}번째 줄: "{이름}" 은(는) 주간 프로그램에 없는 항목이에요')
                continue
            try:
                새값 = ast.literal_eval(값.strip())
            except Exception:
                문제.append(f'지점설정.txt {번호}번째 줄: "{이름}" 값의 모양이 맞지 않아요 (따옴표·괄호를 확인해 주세요)')
                continue
            if 새값 != g[이름]:
                g[이름] = 새값
                바뀜.append(이름)
        break
    if 추가_공휴일 is not None:
        B['추가_공휴일'] = list(g['추가_공휴일'])
        B['_휴일집합'] = None
    return 바뀜, 문제


요일글자 = '월화수목금토일'


# ---------------------------------------------------------------------
# 읽기: 결제내역
# ---------------------------------------------------------------------
def _시각분(s):
    m = re.search(r'(\d{1,2})[-/.](\d{1,2})\s+(\d{1,2}):(\d{2})', str(s or ''))
    if not m:
        return None
    return int(m.group(1)), int(m.group(2)), int(m.group(3)) * 60 + int(m.group(4))


def 결제_읽기(경로):
    ws = 엑셀열기(경로, data_only=True).worksheets[0]
    머리 = None
    열 = {}
    for r in ws.iter_rows(min_row=1, max_row=6):
        글 = {str(c.value).strip(): c.column for c in r if c.value}
        if '대상자명' in 글 and '제공인력명' in 글:
            머리 = r[0].row
            열 = 글
            break
    if not 머리:
        raise SystemExit('결제내역에서 "대상자명 / 제공인력명" 제목 줄을 찾지 못했어요.')

    def 찾기(*이름들):
        for 이름 in 이름들:
            for k, v in 열.items():
                if 이름 in k:
                    return v
        return None
    c = dict(이름=찾기('대상자명'), 등급=찾기('대상자등급'), 인력=찾기('제공인력명'), 구분=찾기('결제구분'), 유형=찾기('서비스유형'),
             시작=찾기('시작시간'), 끝=찾기('종료시간'), 시간=찾기('제공시간', '총시간'), 사유=찾기('일괄결제', '소급사유'), 사업=찾기('사업유형'))
    새형식 = 찾기('제공시간') is None
    자료, 건너뜀 = [], []
    for r in ws.iter_rows(min_row=머리 + 1, values_only=True):
        이름 = r[c['이름'] - 1]
        if not 이름:
            continue
        if c['구분'] and '취소' in str(r[c['구분'] - 1] or ''):          # 취소된 결제 줄은 뺍니다
            continue
        if 사업유형_글자 and c['사업'] and 사업유형_글자 not in str(r[c['사업'] - 1] or ''):
            continue
        s = _시각분(r[c['시작'] - 1])
        e = _시각분(r[c['끝'] - 1])
        if not s:
            건너뜀.append((str(이름), f'결제 시작시간을 읽지 못한 줄이 있어서 뺐어요: {r[c["시작"] - 1]}'))
            continue
        try:
            if 새형식:                                  # '00300' = 3시간 00분 / '00230' = 2시간 30분
                _t = re.sub(r'\D', '', str(r[c['시간'] - 1] or '0')).zfill(5)
                시간 = int(_t[:-2]) + int(_t[-2:]) / 60
            else:
                시간 = float(r[c['시간'] - 1] or 0)
        except Exception:
            시간 = 0.0
        자료.append(dict(아동=str(이름).strip(), 등급=str(r[c['등급'] - 1] or '') if c['등급'] else '', 인력=str(r[c['인력'] - 1] or '').strip(),
                       유형=str(r[c['유형'] - 1] or '').strip(), 월=s[0], 일=s[1], 시작=s[2], 끝=(e[2] if e else s[2]),
                       시간=시간, 사유=str(r[c['사유'] - 1] or '').strip() if c['사유'] else ''))
    return 자료, 건너뜀


# ---------------------------------------------------------------------
# 읽기: 시간표 (시간대 블록 × 주차 칸 × 요일)
# ---------------------------------------------------------------------
class 시간표:
    def __init__(self, 경로, 연도):
        self.ok = False
        self.블록 = []          # [{시: (머리줄, [(시작날짜, 끝날짜, 줄)])}]
        self.ws = None
        self.여러개 = []
        if not 경로:
            return
        try:
            wb = 엑셀열기(경로, data_only=True)
        except Exception:
            return
        ws = wb['시간표'] if '시간표' in wb.sheetnames else wb.worksheets[0]
        self.ws = ws
        self.병합 = {}
        for r in ws.merged_cells.ranges:
            for rr in range(r.min_row, r.max_row + 1):
                for cc in range(r.min_col, r.max_col + 1):
                    self.병합[(rr, cc)] = (r.min_row, r.min_col)
        self.요일열 = {}
        for r in ws.iter_rows(min_row=1, max_row=3):
            for c in r:
                v = str(c.value or '').strip()
                if len(v) == 1 and v in 요일글자[:6] and c.column not in self.요일열:
                    self.요일열[c.column] = 요일글자.index(v)
            if len(self.요일열) >= 4:
                break
        if len(self.요일열) < 4:
            return
        self.열 = {v: k for k, v in self.요일열.items()}
        현재 = None
        마지막시 = 99
        줄들 = list(range(1, ws.max_row + 1))
        머리줄 = []
        for r in 줄들:
            a = str(ws.cell(r, 1).value or '').strip()
            m = re.match(r'(\d{1,2})\s*시\s*~\s*(\d{1,2})\s*시', a)
            if m:
                머리줄.append((r, int(m.group(1)), int(m.group(2))))
            elif 시간표_새모양:                      # '09:00 - 10:00' / '9 - 10시' / '10 - 11:30' 모양 (양천점)
                m = re.match(r'(\d{1,2})(?::(\d{2}))?\s*시?\s*[-~]\s*(\d{1,2})(?::(\d{2}))?\s*시?(?![\d/])', a)
                if m and '/' not in a.split('\n')[0]:
                    머리줄.append((r, int(m.group(1)) + (int(m.group(2)) / 60 if m.group(2) else 0), int(m.group(3)) + (1 if m.group(4) else 0)))
        for i, (r, h, he) in enumerate(머리줄):
            if 현재 is None or h <= 마지막시:
                현재 = {}
                self.블록.append(현재)
            마지막시 = h
            끝줄 = (머리줄[i + 1][0] - 1) if i + 1 < len(머리줄) else ws.max_row
            주들 = []
            for rr in range(r + 1, 끝줄 + 1):
                b = str(ws.cell(rr, 2).value or '').strip()
                mm = re.match(r'(\d{1,2})/(\d{1,2})\s*~\s*(\d{1,2})/(\d{1,2})', b)
                if not mm and 시간표_새모양:           # 'A칸에 9/1 - 4 (1주)' 모양
                    a2 = str(ws.cell(rr, 1).value or '').strip()
                    m2 = re.match(r'(\d{1,2})/(\d{1,2})\s*[-~]\s*(\d{1,2})(?:/(\d{1,2}))?', a2)
                    if m2:
                        _g = (m2.group(1), m2.group(2), m2.group(4) and m2.group(3) or m2.group(1), m2.group(4) or m2.group(3))
                        mm = type('M', (), {'group': staticmethod(lambda n, _g=_g: _g[n - 1])})
                if mm:
                    try:
                        주들.append((dt.date(연도, int(mm.group(1)), int(mm.group(2))), dt.date(연도, int(mm.group(3)), int(mm.group(4))), rr))
                    except ValueError:
                        pass
            for x in range(int(h), max(he, int(h) + 1)):
                현재[x] = (r, 주들)
            if h != int(h):
                현재[h] = (r, 주들)
        self.ok = bool(self.블록)

    def _값(self, r, c):
        k = self.병합.get((r, c), (r, c))
        v = self.ws.cell(*k).value
        return re.sub(r'\s+', ' ', str(v)).strip() if v is not None else ''

    def 칸(self, 블록번호, 날짜, 시):
        """(위칸 글, 아래(주차)칸 글) — 없으면 ('', '')"""
        if not self.ok or 블록번호 >= len(self.블록):
            return '', ''
        항목 = self.블록[블록번호].get(시)
        c = self.열.get(날짜.weekday())
        if not 항목 or not c:
            return '', ''
        머리, 주들 = 항목
        위 = self._값(머리, c)
        if not 위 and 시간표_새모양:                    # 머리줄에 칸이 하나만 적혀 있으면 모든 요일 공통으로 봐요 (예: 11:30~13시 생활훈련)
            채움 = [self._값(머리, cc) for cc in self.열.values()]
            채움 = [x for x in 채움 if x and x != '-']
            if len(set(채움)) == 1 and len(채움) == 1:
                위 = 채움[0]
        아래 = ''
        for a, b, rr in 주들:
            if a <= 날짜 <= b:
                아래 = self._값(rr, c)
                break
        if 위 == '-':
            위 = ''
        if 아래 == '-':
            아래 = ''
        return self._하나만(위, 날짜), self._하나만(아래, 날짜)

    @staticmethod
    def _나누기(글):
        조각, 깊이, 지금 = [], 0, ''
        for ch in 글:
            if ch in '([':
                깊이 += 1
            elif ch in ')]':
                깊이 -= 1
            if ch == '/' and 깊이 == 0:
                조각.append(지금.strip())
                지금 = ''
            else:
                지금 += ch
        조각.append(지금.strip())
        return [x for x in 조각 if x]

    def _하나만(self, 글, 날짜):
        """'볼링/네일아트/감각놀이'처럼 / 로 묶인 활동은 날짜마다 돌아가며 하나만 써요 (같은 날은 모두 같은 것)"""
        조각 = self._나누기(글)
        if len(조각) <= 1 or not 슬래시_고르기:
            return 글
        return 조각[(날짜.toordinal() // 7 + 날짜.weekday()) % len(조각)]


def 칸_글(위, 아래):
    if 소제목_줄 and 위 and 아래 and 위 != 아래:        # 대제목 줄 + ' -소제목' 줄 (양천 최중증)
        첫 = [x.strip() for x in re.split(r'\s*[&/]\s*', 위) if x.strip()]
        if any(아래.startswith(x) or x.startswith(아래.split('(')[0].strip()) for x in 첫 if x):
            return 아래                                   # 소제목이 이미 대제목 이름으로 시작하면 소제목만
        return f'{위}\n -' + re.sub(r'\s+\(', '(', 아래)
    if 상세로_대체 and 아래:
        return 아래
    if not 아래 or 아래 == 위:
        return 위 or 아래
    if not 위:
        return 아래
    if '(협)' in 아래 or 위 in 상위이름_생략:
        return 아래
    앞 = 위.split('(')[0]
    if 앞 and 아래.startswith(앞):
        return 아래
    if 위.endswith(')') and '(' in 위:
        return 위[:-1] + '/' + 아래 + ')'
    return f'{위}({아래})'


def 종류(글):
    if 글.startswith(송영_문구):
        return '송'
    if any(g in 글 for g in 협력_글자):
        return '협'
    if any(g in 글 for g in 외부활동_글자):
        return '외'
    return '내'


def _유효(y, m, d):
    try:
        dt.date(y, m, d)
        return True
    except ValueError:
        return False


def 시각글(t):
    h = int(t + 1e-9)
    m = 30 if (t - h) >= 0.5 - 1e-9 else 0
    if 시각_표기 == '콜론':
        return f'{h:02d}:{m:02d}'
    return f'{h}시' if m == 0 else f'{h}시{m}분'


def 숫자(x):
    return int(x) if abs(x - round(x)) < 1e-9 else x


def 시간글자(h):
    return f'{h:g}'


def 한줄날짜(d):
    return f'{d.year}.{d.month:02d}.{d.day:02d}.({요일글자[d.weekday()]})'


# ---------------------------------------------------------------------
# 하루 만들기
# ---------------------------------------------------------------------
def 보조고르기(아동, 날짜, 시):
    n = len(보조_활동)
    return 보조_활동[(sum(ord(ch) for ch in 아동) + 날짜.toordinal() + int(시 * 3)) % n]


def _시작시(분):
    return (round(분 / 30) * 30 / 60) if 고정_칸 else ((분 + 30) // 60)


def 하루_칸들(아동, 날짜, 시작분, 시간, 송영있음, 묶음번호, tt):
    """그날의 시간 칸 목록(1시간씩)을 만들어요. 항목: dict(a,b,위,아래,글,종류,출처)"""
    칸들 = []
    CAP = 하루_마지막_시각
    P0 = (시작분 + 30) // 60
    if 송영있음:
        T = min(P0 - 1 if 송영_시간에_포함 else P0, int(CAP - 시간 - 1))
        T = max(T, 7)
        칸들.append(dict(a=T, b=T + 1, 위=송영_문구, 아래='', 글=송영_문구, 종류='송', 출처='송영'))
        P = T + 1
    else:
        P = max(7, min(P0, int(CAP - math.ceil(시간))))
    끝 = P + 시간
    t = P
    if 고정_칸:                                    # 칸 경계가 정해져 있는 곳(최중증: 9:30 10 11:30 13 14 15) — 결제 시각으로 앞뒤를 잘라 써요
        t0 = round(시작분 / 30) * 30 / 60
        끝 = t0 + 시간
        구간들 = [(max(e[0], t0), min(e[1], 끝), (e[2] if len(e) > 2 else e[0])) for e in 고정_칸 if e[1] > t0 + 1e-9 and e[0] < 끝 - 1e-9]
        if 구간들:
            t = 구간들[0][0]
        else:                                      # 정해진 칸 밖 시각이면 1시간 칸으로 평소처럼
            구간들 = None
            끝 = P + 시간
            t = P
    else:
        구간들 = None
    _순번 = 0
    while t < 끝 - 1e-9:
        if 구간들 is not None:
            t, b, h = 구간들[_순번]
            _순번 += 1
        else:
            h = int(t + 1e-9)
            b = min(h + 1, 끝)
        위, 아래 = tt.칸(묶음번호, 날짜, h) if tt else ('', '')
        if tt and not (위 or 아래) and h != int(h):       # 30분 칸은 같은 시각의 큰 칸(예: 9~10시)에서 찾아봐요
            위, 아래 = tt.칸(묶음번호, 날짜, int(h))
        if '연휴' in 위 or '연휴' in 아래:
            위 = 아래 = ''
        글 = 칸_글(위, 아래)
        if 괄호붙이기:
            글 = re.sub(r'\s+\(', '(', 글)
        for a_, b_ in 글자_바꾸기.items():
            글 = 글.replace(a_, b_)
        출처 = '표'
        if 글 == '자율활동':
            후보 = [x for x in 보조_활동 if x.startswith('자율활동')]
            if 후보:
                글 = 후보[(sum(ord(ch) for ch in 아동) + 날짜.toordinal() + int(h * 3)) % len(후보)]
                위, 아래 = 글, ''
        if not 글:
            출처 = '보조'
            if h >= 16 and b >= 끝 - 1e-9 and 보조_외부활동 and (끝 >= 17):
                글 = 보조_외부활동
            else:
                글 = 보조고르기(아동, 날짜, h)
            위, 아래 = 글, ''
        elif 점심_문구 and h == 12 and '식사' in 글:
            글 = 점심_문구
            위, 아래 = 글, ''
        칸들.append(dict(a=t, b=b, 위=위, 아래=아래, 글=글, 종류=종류(글), 출처=출처))
        t = b
        if 구간들 is not None and _순번 >= len(구간들):
            break
    return 칸들


def 합치기(칸들):
    묶음 = []
    for s in sorted(칸들, key=lambda x: x['a']):
        if 묶음:
            p = 묶음[-1]
            같다 = (p['종류'] == s['종류'] and p['출처'] == s['출처'] and abs(p['b'] - s['a']) < 1e-9
                  and ((s['출처'] == '표' and p['위'] and p['위'] == s['위']) or (s['출처'] != '표' and p['글'] == s['글'])))
            if 같다 and s['종류'] != '송':
                p['b'] = s['b']
                if not p['아래'] and s['아래']:
                    p['아래'] = s['아래']
                    p['글'] = 칸_글(p['위'], p['아래'])
                    p['종류'] = 종류(p['글'])
                continue
        묶음.append(dict(s))
    return 묶음


def _특이풀기(글):
    return [x.strip(' -') for x in 글.split('\n') if x.strip(' -')] if 글 else []


def _특이합(목록):
    if 특이_줄:
        return '\n'.join(' -' + x.strip() for x in 목록 if x and x.strip())
    return ' / '.join(목록)


def 글_만들기(칸들):
    return 칸_구분.join(f"{시각글(s['a'])}~{시각글(s['b'])} {s['글']}" for s in 합치기(칸들))


def 시간합(칸들):
    r = {'내': 0.0, '외': 0.0, '협': 0.0, '송': 0.0}
    for s in 칸들:
        r[s['종류']] += s['b'] - s['a']
    return r


# ---------------------------------------------------------------------
# 시트 채우기
# ---------------------------------------------------------------------
def 시트_채우기(ws, 정보, 계획서):
    월 = 정보['월']
    # 양식 위쪽 빈 줄 수가 달라도 맞도록: '날짜' 머리줄(기본 11줄) 위치로 밀림 계산
    오 = 0
    for _r in range(1, 25):
        if str(ws.cell(_r, 1).value or '').replace('\n', '').replace(' ', '') == '날짜':
            오 = _r - 11
            break
    def 칸(주소):
        m_ = re.match(r'([A-Z]+)(\d+)$', 주소)
        return f'{m_.group(1)}{int(m_.group(2)) + 오}'
    ws[칸('A3')] = f'주간활동서비스 ({월})월   ■활동계획서   □활동기록지' if 계획서 else f'주간활동서비스 ({월})월   □활동계획서   ■활동기록지'
    ws[칸('D4')] = 정보['담당']
    ws[칸('J4')] = 한줄날짜(정보['계획일'] if 계획서 else 정보['기록일'])
    ws[칸('D5')] = 정보['아동']
    e6 = 정보['담당'] + (f" / {정보['송영담당']}(송영)" if 정보['송영담당'] and 정보['송영담당'] != 정보['담당'] else '')
    ws[칸('E6')] = e6
    ws[칸('E6')].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True, shrink_to_fit=True)
    ws[칸('G6')] = ', '.join(정보['대체인력'])
    ws[칸('G6')].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True, shrink_to_fit=True)
    수급 = 정보['수급']
    ws[칸('D7')] = ('■ 월 132시간    □ 월 176시간 ' if 수급 == 수급시간_기본형 else '□ 월 132시간    ■ 월 176시간 ') \
        if (수급시간_기본형, 수급시간_확장형) == (132, 176) else \
        (f'■ 월 {수급}시간' )
    ws[칸('D8')] = f"월 ({시간글자(정보['계획총'])})시간"
    ws[칸('D9')] = 정보['축소사유']
    행 = 12 + 오
    마지막 = 행 - 1
    for 일 in 정보['날들']:
        if 행 > 34 + 오:
            break
        ws.cell(행, 1).value = 일['날짜'].day
        ws.cell(행, 2).value = 요일글자[일['날짜'].weekday()]
        ws.cell(행, 3).value = 일['인력']
        ws.cell(행, 4).value = 일['글계획']
        시 = 일['시간계획'] if 계획서 else 일['시간']
        if not 계획서:
            ws.cell(행, 7).value = 일['글']
            ws.cell(행, 10).value = 일['특이'] or None
        else:
            ws.cell(행, 10).value = None
        for 열번호, 가로 in ((4, 'left'), (7, 'left'), (10, 'center')):
            ws.cell(행, 열번호).alignment = Alignment(horizontal=가로, vertical='center', wrap_text=True)
        ws.cell(행, 12).value = 숫자(시['내'])
        ws.cell(행, 13).value = 숫자(시['외'])
        ws.cell(행, 14).value = 숫자(시['협'])
        ws.cell(행, 15).value = 숫자(시['송'])
        줄 = max(줄수(일['글계획'], 46), 줄수(일['글'], 46) if not 계획서 else 1, 줄수(일['특이'] if not 계획서 else '', 22))
        ws.row_dimensions[행].height = max(39.95, 15.0 * 줄 + 4)
        마지막 = 행
        행 += 1
    # 요구사항 줄 (날짜 줄 바로 아래)
    r = 마지막 + 1
    if r <= 34 + 오 and not (빈_요구사항_줄_빼기 and not 정보['요구사항'] and not (인정결제_요구사항에 and 정보['인정시간'] and not 계획서)):
        글 = 요구사항_머리말 + (정보['요구사항'] or '')
        if (not 계획서) and 정보['인정시간'] and 인정결제_요구사항에:
            글 = 글.rstrip() + f" / 인정결제 {시간글자(정보['인정시간'])}시간"
        for rng in list(ws.merged_cells.ranges):
            if rng.min_row == r and rng.max_row == r:
                ws.unmerge_cells(str(rng))
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=15)
        c = ws.cell(r, 1)
        c.value = 글
        c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        c.font = copy_font(ws.cell(12 + 오, 4).font)
        ws.row_dimensions[r].height = max(30.0, 15.0 * 줄수(글, 100) + 6)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    기준일 = 정보['계획일'] if 계획서 else 정보['기록일']
    ws[칸('A40')] = f'     {기준일.year}년         {기준일.month:02d}월         {기준일.day}일       '
    old = str(ws[칸('A42')].value or '')
    ws[칸('A42')] = re.sub(r'(보호자\)명\s+)\S+(\s+\(인\))', lambda m: m.group(1) + 정보['아동'] + m.group(2), old) if re.search(r'보호자\)명', old) \
        else f'이용자(혹은  보호자)명                          {정보["아동"]}   (인)      '
    if 정보['아동'] not in str(ws[칸('A42')].value):
        ws[칸('A42')] = f'이용자(혹은  보호자)명                          {정보["아동"]}   (인)      '


def copy_font(f):
    from copy import copy
    return copy(f)


# ---------------------------------------------------------------------
# 파일 찾기
# ---------------------------------------------------------------------
def 파일_고르기(종류_, 월=None):
    return B['파일_고르기']('', 종류_, 월)


# ---------------------------------------------------------------------
# 실행
# ---------------------------------------------------------------------
def main():
    global 연도
    바뀜, 설정문제 = 설정_불러오기()
    if 바뀜:
        print('지점설정.txt 를 적용했어요:', ', '.join(바뀜))
    for 글 in 설정문제:
        print('※', 글)
    결제경로 = 파일_고르기('결제')
    if not 결제경로:
        raise SystemExit('결제내역 엑셀을 찾지 못했어요.')
    자료, 건너뜀 = 결제_읽기(결제경로)
    if not 자료:
        raise SystemExit('결제내역에서 읽을 내용이 없어요.')
    월 = Counter(r['월'] for r in 자료).most_common(1)[0][0]
    if not 연도:
        오늘 = dt.date.today()
        연도 = 오늘.year - 1 if 월 > 오늘.month else 오늘.year
    점검 = [(a, None, 글) for a, 글 in 건너뜀]
    다른달 = Counter(r['월'] for r in 자료 if r['월'] != 월)
    for 달, 개수 in sorted(다른달.items()):
        점검.append(('(결제내역 전체)', None, f'{달}월 결제 {개수}줄은 이번 달({월}월)이 아니라서 뺐어요'))
    자료 = [r for r in 자료 if r['월'] == 월]
    for r in 자료:
        r['날짜'] = dt.date(연도, r['월'], r['일'])

    양식경로 = 파일_고르기('양식')
    시간표경로 = 파일_고르기('시간표', 월)
    if not 양식경로:
        raise SystemExit('주간활동서비스 빈 양식 엑셀을 찾지 못했어요. 첫 줄에 "활동계획서 / 활동기록지"가 있는 양식 파일을 올려 주세요.')
    print('결제내역:', os.path.basename(결제경로))
    print('시간표  :', os.path.basename(시간표경로) if 시간표경로 else '(없음)')
    print('양식    :', os.path.basename(양식경로))
    tt = 시간표(시간표경로, 연도)
    if not 시간표경로:
        print('※ 시간표 파일을 못 찾았어요. 시간표 없이 기본 활동으로만 채워요.')
        점검.append(('(전체)', None, '시간표 파일이 없어서 활동이 모두 기본 활동(자율활동 등)으로 채워졌어요'))
    elif not tt.ok:
        print('※ 시간표 모양을 읽지 못했어요.')
        점검.append(('(전체)', None, '시간표 모양을 읽지 못해서 활동이 모두 기본 활동으로 채워졌어요'))
    else:
        print(f'시간표 읽음: 표 {len(tt.블록)}개')
        if len(tt.블록) >= 2:
            print('  · 위쪽 표는 9시 시작, 아래쪽 표는 10시 시작(송영 다음)으로 보고 있어요')
        여러 = set()
        for bi, 블 in enumerate(tt.블록):
            for h, (머리, 주들) in 블.items():
                for c in tt.열.values():
                    for _, _, rr in 주들:
                        v = tt._값(rr, c)
                        if v.count('/') >= 2 and '(협)' in v:
                            여러.add((v, h))
        for v, h in sorted(여러):
            점검.append(('(시간표)', None, f'{h}시 칸에 여러 활동이 / 로 묶여 있어요: "{v}". / 로 묶인 활동은 날짜마다 돌아가며 하나만 넣었어요'))

    wb0 = 엑셀열기(양식경로)
    원본양식 = wb0.worksheets[0]
    전체 = wb0

    아동별 = defaultdict(list)
    for r in 자료:
        아동별[r['아동']].append(r)
    대상 = sorted(a for a, rs in 아동별.items() if any(x['유형'] == '주간활동서비스' for x in rs))
    if not 대상:
        raise SystemExit('결제내역에서 "주간활동서비스" 결제를 찾지 못했어요.')

    바탕 = None
    for 후보 in (os.path.join(os.path.expanduser('~'), 'Desktop'), os.path.join(os.path.expanduser('~'), '바탕 화면')):
        if os.path.isdir(후보):
            바탕 = 후보
            break
    폴더이름 = f'{출력_폴더}_{월}월' + (f'_{지점이름}' if 지점이름 else '')
    출력 = os.path.join(바탕 or os.getcwd(), 폴더이름)
    os.makedirs(출력, exist_ok=True)

    요약 = []
    계획일 = B['전달_마지막_근무일'](연도, 월)
    사유없음, 시간다름 = [], defaultdict(list)
    시간표활동바꿈 = defaultdict(float)
    센터운영일 = {r['날짜'] for r in 자료 if r['유형'] == '주간활동서비스'}
    월간_해석 = []
    for 항 in 월간_활동:
        try:
            if 항.get('날짜'):                                  # 특정 날짜 하루만 (예: '10/22')
                _m = re.match(r'\s*(\d{1,2})\s*[/.\-]\s*(\d{1,2})', str(항['날짜']))
                대상날 = dt.date(연도, int(_m.group(1)), int(_m.group(2)))
                요일 = 대상날.weekday()
            else:
                요일 = 항['요일'] if isinstance(항['요일'], int) else 요일글자.index(str(항['요일'])[0])
                번째 = int(항.get('주차', 2))
                날짜들_ = [dt.date(연도, 월, d_) for d_ in range(1, 32) if _유효(연도, 월, d_) and dt.date(연도, 월, d_).weekday() == 요일]
                대상날 = 날짜들_[min(번째, len(날짜들_)) - 1]
            이용 = list(항.get('이용자') or [])
            if not 이용 and 항.get('날짜') and not 항.get('인원'):
                이용 = sorted(대상)                                # 날짜만 적고 이용자를 비우면 모두
            if not 이용:
                n = int(항.get('인원', 0))
                이용 = sorted(대상, key=lambda a_: sum(ord(c_) * (i_ + 3) for i_, c_ in enumerate(a_)))[:n]
                점검.append(('(전체)', None, f'{항["이름"]} 참여자 {len(이용)}명({", ".join(이용)})을 알아서 골랐어요. 정해진 사람이 있으면 설정의 월간_활동에 이용자 이름을 적어 주세요'))
            월간_해석.append(dict(이름=항['이름'], 날짜=대상날, 시작=항['시작'], 끝=항.get('끝'), 이용자=set(이용)))
            점검.append(('(전체)', None, f'{항["이름"]}: {대상날.month}/{대상날.day}({요일글자[대상날.weekday()]}) {항["시작"]}시에 넣었어요'))
        except Exception as e_:
            점검.append(('(전체)', None, f'월간_활동 설정을 읽지 못했어요 ({e_})'))
    for 아동 in 대상:
        rs = 아동별[아동]
        주간 = defaultdict(list)
        송영 = defaultdict(list)
        인정 = []
        for x in rs:
            if x['유형'] == '주간활동서비스':
                주간[x['날짜']].append(x)
            elif x['유형'] == '송영서비스':
                if not 송영_대상자 or 아동 in 송영_대상자:
                    송영[x['날짜']].append(x)
            elif x['유형'] == '인정결제':
                인정.append(x)
        등급 = rs[0]['등급']
        수급 = 수급시간_확장형 if '확장' in 등급 else 수급시간_기본형
        if '기본' not in 등급 and '확장' not in 등급:
            점검.append((아동, None, f'대상자등급 "{등급}" 에서 기본/확장을 알 수 없어서 기본형 {수급시간_기본형}시간으로 봤어요'))
        주인 = Counter(x['인력'] for d in 주간.values() for x in d)
        담당 = 담당_지정.get(아동) or (주인.most_common(1)[0][0] if 주인 else '')
        송영주인 = Counter(x['인력'] for d in 송영.values() for x in d)
        송영담당 = 송영주인.most_common(1)[0][0] if 송영주인 else ''
        시간들 = [sum(x['시간'] for x in d) for d in 주간.values()]
        보통 = Counter(시간들).most_common(1)[0][0] if 시간들 else 0
        시작들 = [_시작시(min(x['시작'] for x in d)) for d in 주간.values()]
        보통시작 = Counter(시작들).most_common(1)[0][0] if 시작들 else 9

        날들 = []
        칸모음 = {}
        날짜들 = sorted(set(주간) | set(송영))
        for 날짜 in 날짜들:
            if 날짜.weekday() >= 5:
                점검.append((아동, 날짜, '토·일요일 결제가 있어요. 주간활동서비스는 평일만 운영한다고 하셔서 이 날은 뺐어요'))
                continue
            d주 = 주간.get(날짜, [])
            시간 = sum(x['시간'] for x in d주)
            시작분 = min(x['시작'] for x in d주) if d주 else min(x['시작'] for x in 송영[날짜])
            있음송영 = 날짜 in 송영
            묶음번호 = 0
            if 아동 in 묶음_지정 and str(묶음_지정[아동]).upper() == 'B' and len(tt.블록) >= 2:
                묶음번호 = 1
            elif 있음송영 and 송영_있는날_두번째표 and len(tt.블록) >= 2:
                묶음번호 = 1
            T_ = tt if tt.ok else None
            차 = (보통 - 시간) if (d주 and 보통) else 0
            if d주 and 보통:
                계획칸 = 하루_칸들(아동, 날짜, 보통시작 * 60, 보통, 있음송영, 묶음번호, T_)
            else:
                계획칸 = 하루_칸들(아동, 날짜, 시작분, 시간, 있음송영, 묶음번호, T_)
            늦음 = _시작시(시작분) > 보통시작
            if 차 > 1e-9:
                프로그램 = [s for s in 계획칸 if s['종류'] != '송']
                ps, pe = min(s['a'] for s in 프로그램), max(s['b'] for s in 프로그램)
                ra, rb = (ps + 차, pe) if 늦음 else (ps, pe - 차)
                기록칸 = []
                for s in 계획칸:
                    if s['종류'] == '송':
                        기록칸.append(s)
                    elif s['b'] <= ra + 1e-9 or s['a'] >= rb - 1e-9:
                        continue
                    elif s['a'] >= ra - 1e-9 and s['b'] <= rb + 1e-9:
                        기록칸.append(s)
                    else:
                        c = dict(s)
                        c['a'], c['b'] = max(s['a'], ra), min(s['b'], rb)
                        c['orig'] = s
                        기록칸.append(c)
            elif 차 < -1e-9:
                새 = 하루_칸들(아동, 날짜, 시작분, 시간, 있음송영, 묶음번호, T_)
                기록칸 = []
                for n in 새:
                    같은 = [o for o in 계획칸 if abs(o['a'] - n['a']) < 1e-9 and abs(o['b'] - n['b']) < 1e-9]
                    기록칸.append(같은[0] if 같은 else n)
            else:
                기록칸 = 계획칸
            for 월간 in 월간_해석:
                if 월간['날짜'] == 날짜 and 아동 in 월간['이용자'] and 월간.get('끝') and 월간['끝'] - 월간['시작'] > 1 + 1e-9:
                    for 대상칸 in (계획칸, 기록칸):                   # 여러 시간짜리 (예: 9~15시 밤따기) → 그 시간대를 한 칸으로 바꿔요
                        프로 = [s_ for s_ in 대상칸 if s_['종류'] != '송']
                        if not 프로:
                            continue
                        가 = max(월간['시작'], min(s_['a'] for s_ in 프로))
                        나 = min(월간['끝'], max(s_['b'] for s_ in 프로))
                        if 나 - 가 <= 1e-9:
                            continue
                        새칸 = []
                        for s_ in 대상칸:
                            if s_['종류'] == '송' or s_['b'] <= 가 + 1e-9 or s_['a'] >= 나 - 1e-9:
                                새칸.append(s_)
                            else:
                                if s_['a'] < 가 - 1e-9:
                                    c_ = dict(s_); c_['b'] = 가; 새칸.append(c_)
                                if s_['b'] > 나 + 1e-9:
                                    c_ = dict(s_); c_['a'] = 나; 새칸.append(c_)
                        글_ = 월간['이름']
                        새칸.append(dict(a=가, b=나, 위=글_, 아래='', 글=글_, 종류=종류(글_) if 종류(글_) != '내' else '외', 출처='특별'))
                        대상칸[:] = sorted(새칸, key=lambda x_: x_['a'])
                    continue
                if 월간['날짜'] == 날짜 and 아동 in 월간['이용자']:
                    for 대상칸 in (계획칸, 기록칸):
                        for s_ in 대상칸:
                            if s_['종류'] != '송' and abs(s_['a'] - 월간['시작']) < 1e-9:
                                글_ = 월간['이름']
                                s_.update(위=글_, 아래='', 글=글_, 종류=종류(글_) if 종류(글_) != '내' else '외', 출처='특별')
            칸모음[날짜] = (계획칸, 기록칸)
            인력 = (d주[0]['인력'] if d주 else 송영[날짜][0]['인력'])
            특이 = []
            사유 = []
            for x in d주:
                if x['사유'] and 소급_문구.get(x['사유']):
                    사유.append(소급_문구[x['사유']])
                elif x['사유']:
                    사유.append(x['사유'])
            for x in 송영.get(날짜, []):
                if x['사유']:
                    문 = 소급_문구.get(x['사유'], x['사유'])
                    사유.append('송영서비스 ' + 문)
            if 인력 != 담당 and 담당:
                특이.append(대체인력_문구.replace('{담당}', 담당).replace('{대체}', 인력))
            if 있음송영 and 송영담당:
                for x in 송영[날짜]:
                    if x['인력'] and x['인력'] != 송영담당 and 송영주인.get(송영담당, 0) >= 3:
                        특이.append(f'송영서비스 제공인력 변경({x["인력"]})')
                        break
            특이 += list(dict.fromkeys(사유))
            if 시간변동_표기 and d주 and 보통 and abs(차) > 1e-9:
                if 차 < 0:
                    시간다름[아동].append(f'{날짜.month}/{날짜.day}(+{시간글자(-차)}시간 추가)')
                    사 = 추가이용_사유 if isinstance(추가이용_사유, (list, tuple)) else [추가이용_사유]
                    특이.append(f'{시간글자(-차)}시간 추가이용 (사유:{사[(날짜.toordinal() // 3) % len(사)] if 사 else ""})')
                else:
                    시간다름[아동].append(f'{날짜.month}/{날짜.day}(-{시간글자(차)}시간 {"늦은 등원" if 늦음 else "이른 하원"})')
                    특이.append(f'{시간글자(차)}시간 {"늦은 등원" if 늦음 else "이른 하원"} (사유:{늦은등원_사유 if 늦음 else 이른하원_사유})')
            날들.append(dict(날짜=날짜, 인력=인력, 특이=_특이합(특이), 시간=None, 글=None))

        # ---- 결제가 없는 평일: 센터 전체가 쉬었으면 휴일, 아니면 불참
        for d_ in range(1, 32):
            if not _유효(연도, 월, d_):
                continue
            날짜 = dt.date(연도, 월, d_)
            if 날짜.weekday() >= 5 or 날짜 in 주간 or 날짜 in 송영 or (not 인정결제_요구사항에 and any(x_['날짜'] == 날짜 for x_ in 인정)):
                continue
            if 날짜 not in 센터운영일:
                이름_ = ''
                for h_ in range(9, 17):
                    for b_ in range(len(tt.블록) if tt.ok else 0):
                        위_, 아래_ = tt.칸(b_, 날짜, h_)
                        for t_ in (위_, 아래_):
                            if '연휴' in t_:
                                이름_ = t_
                if not 이름_:
                    이름_ = '공휴일' if B['쉬는날'](날짜) else '휴무'
                continue   # 센터가 쉰 날(연휴 등)은 줄을 만들지 않아요
            elif 보통:
                계획칸 = 하루_칸들(아동, 날짜, 보통시작 * 60, 보통, False, 1 if (아동 in 묶음_지정 and str(묶음_지정[아동]).upper() == 'B' and len(tt.블록) >= 2) else 0, tt if tt.ok else None)
                칸모음[날짜] = (계획칸, [])
                사_ = 불참_사유 if isinstance(불참_사유, (list, tuple)) else [불참_사유]
                _지정 = (불참_지정.get(아동) or {}) if isinstance(불참_지정, dict) else {}
                _직접 = _지정.get(f'{날짜.month}/{날짜.day}') or _지정.get(f'{날짜.month:02d}/{날짜.day:02d}')
                _사유 = _직접 if _직접 else (사_[(sum(ord(c) for c in 아동) + 날짜.toordinal()) % len(사_)] if 사_ else '')
                문 = 불참_문구.replace('{사유}', _사유)
                날들.append(dict(날짜=날짜, 인력=담당, 특이=문, 시간=None, 글=None, 글계획=None, 시간계획=None, 불참=True))
        if not 인정결제_요구사항에:                      # 인정결제는 그 날짜 줄의 개별 특이사항에 적어요 (활동 없는 날이면 줄을 하나 더 만들어요)
            인정합 = defaultdict(float)
            인정인력 = {}
            for x_ in 인정:
                인정합[x_['날짜']] += x_['시간']
                인정인력.setdefault(x_['날짜'], x_.get('인력') or 담당)
            for 날짜_, 시간_ in sorted(인정합.items()):
                글_ = 인정결제_문구.replace('{n}', 시간글자(시간_))
                같은 = [일_ for 일_ in 날들 if 일_['날짜'] == 날짜_]
                if 같은:
                    같은[0]['특이'] = _특이합([x_ for x_ in (같은[0]['특이'], 글_) if x_]) if not 특이_줄 else _특이합(_특이풀기(같은[0]['특이']) + [글_])
                    같은[0]['인정글'] = 글_
                elif 날짜_.weekday() >= 5:
                    점검.append((아동, 날짜_, f'인정결제({글_})가 토·일요일이라 줄을 만들지 못했어요. 확인해 주세요'))
                else:
                    날들.append(dict(날짜=날짜_, 인력=인정인력[날짜_], 특이=_특이합([글_]), 인정글=글_, 시간=None, 글=None, 글계획=None, 시간계획=None, 인정만=True))
        날들.sort(key=lambda x: x['날짜'])

        # ---- 외부(+협력) 30% 맞추기
        def 합_of(번호):
            return sum(s['b'] - s['a'] for 칸 in 칸모음.values() for s in 칸[번호] if s['종류'] in ('외', '협'))

        def 외부합():
            return min(합_of(0), 합_of(1))

        def 모두(날짜):
            본, 목록 = set(), []
            for 칸 in 칸모음[날짜]:
                for s in 칸:
                    if id(s) not in 본:
                        본.add(id(s))
                        목록.append(s)
            return 목록

        def 바꾸기(s, 목록):
            for o in 목록:
                if o is s or o.get('orig') is s or (s.get('orig') is not None and o is s['orig']):
                    o['위'] = 외부_채울_활동
                    o['아래'] = ''
                    o['글'] = 외부_채울_활동
                    o['종류'] = '외'
                    o['출처'] = '전환'
        목표 = 외부_최소비율 * 수급
        처음외부 = 외부합()
        if 처음외부 + 1e-9 < 목표:
            # 하루에 한 시간씩, 여러 날에 골고루 바꿔요 (한 날에 몰리지 않게). 모자라면 한 바퀴 더 돌아요
            def 후보순위(s, 칸, 허용2=False):
                if s['종류'] != '내' or '식사' in s['글'] or '양치' in s['글']:
                    return None
                if s['출처'] == '보조':
                    순위 = 0
                elif any(g in s['글'] for g in 바꿔도_되는_글자):
                    순위 = 1
                elif 허용2 and s['출처'] == '표':
                    순위 = 2
                else:
                    return None
                붙음 = any(o is not s and o['종류'] in ('외', '협') and (abs(o['b'] - s['a']) < 1e-9 or abs(o['a'] - s['b']) < 1e-9) for o in 칸)
                return (순위, 1 if 붙음 else 0)
            두번째 = False
            for 바퀴 in range(8):
                바뀜있음 = False
                for 날짜 in sorted(칸모음):
                    if 외부합() + 1e-9 >= 목표:
                        break
                    칸 = 모두(날짜)
                    if sum(1 for s in 칸 if s['출처'] == '전환') > 바퀴:
                        continue
                    고름 = [(후보순위(s, 칸, 두번째), s['a'], s) for s in 칸]
                    고름 = [x for x in 고름 if x[0] is not None]
                    if not 고름:
                        continue
                    고름.sort(key=lambda x: (x[0], -x[1]))
                    s = 고름[0][2]
                    바꾸기(s, 칸)
                    if 고름[0][0][0] == 2:
                        시간표활동바꿈[아동] += s['b'] - s['a']
                    바뀜있음 = True
                if 외부합() + 1e-9 >= 목표:
                    break
                if not 바뀜있음 and not 두번째:
                    두번째 = True
        외부 = 외부합()
        if 외부 + 1e-9 < 목표:
            점검.append((아동, None, f'외부+협력이 {시간글자(외부)}시간으로 수급시간 {수급}시간의 30%({목표:.1f}시간)에 못 미쳐요. 바꿀 수 있는 활동이 부족해요'))
        elif 외부 > 처음외부 + 1e-9:
            점검.append((아동, None, f'외부+협력이 모자라서(시간표대로면 {시간글자(처음외부)}시간) 자율활동 등 {시간글자(외부 - 처음외부)}시간을 산책으로 바꿨어요'))
            if 시간표활동바꿈[아동]:
                점검.append((아동, None, f'자율활동으로 바꿀 시간이 모자라서 시간표의 다른 활동 {시간글자(시간표활동바꿈[아동])}시간도 산책으로 바꿨어요. 시간표를 확인해 주세요'))

        총내부 = 총외부 = 총협 = 총송 = 0.0
        계획총 = 0.0
        for 일 in 날들:
            if 일.get('인정만'):
                일['글계획'] = 일['글'] = ''
                일['시간'] = 일['시간계획'] = {'내': 0.0, '외': 0.0, '협': 0.0, '송': 0.0}
                continue
            if 일.get('휴일'):
                일['시간'] = 일['시간계획'] = {'내': 0.0, '외': 0.0, '협': 0.0, '송': 0.0}
                continue
            계획칸, 기록칸 = 칸모음[일['날짜']]
            일['글계획'] = 글_만들기(계획칸)
            일['시간계획'] = 시간합(계획칸)
            계획총 += sum(일['시간계획'].values())
            일['글'] = 글_만들기(기록칸) if 기록칸 else ''
            일['시간'] = 시간합(기록칸)
            총내부 += 일['시간']['내']
            총외부 += 일['시간']['외']
            총협 += 일['시간']['협']
            총송 += 일['시간']['송']
        총 = 총내부 + 총외부 + 총협 + 총송

        인정날 = sorted({x['날짜'] for x in 인정})
        인정시간 = sum(x['시간'] for x in 인정)
        if 기록일_방식 == '인정결제' and 인정날:
            기록일 = max(인정날)
        else:
            기록일 = B['이번달_마지막_근무일'](연도, 월)
            if 기록일_방식 == '인정결제' and not 인정날:
                점검.append((아동, None, f'인정결제가 없어서 기록지 작성일자를 {기록일.month}/{기록일.day}로 넣었어요'))
        사유 = 축소사유_지정.get(아동) if 아동 in 축소사유_지정 else _돌려쓰기_고르기(아동)
        if 총 + 1e-9 < 수급 and not 사유:
            사유없음.append(아동)
        대체 = sorted({일['인력'] for 일 in 날들 if 일['인력'] != 담당})
        정보 = dict(아동=아동, 월=월, 담당=담당, 송영담당=송영담당 if 송영 else '', 날들=날들, 총=총, 계획총=계획총, 수급=수급, 계획일=계획일, 기록일=기록일,
                  대체인력=대체, 요구사항='', 인정시간=인정시간, 축소사유=사유)
        계획 = 전체.copy_worksheet(원본양식)
        계획.title = f'{아동}{시트이름_구분}계획서'
        기록 = 전체.copy_worksheet(원본양식)
        기록.title = f'{아동}{시트이름_구분}기록지'
        시트_채우기(계획, 정보, True)
        시트_채우기(기록, 정보, False)
        if 총 + 인정시간 < 수급 - 1e-9 or 총 + 인정시간 > 수급 + 1e-9:
            점검.append((아동, None, f'계획시간 {시간글자(총)} + 인정결제 {시간글자(인정시간)} = {시간글자(총 + 인정시간)}시간이라 수급시간 {수급}시간과 달라요'))
        요약.append((아동, 담당, len([d_ for d_ in 날들 if not d_.get('휴일') and not d_.get('불참') and not d_.get('인정만')]), len(인정날), 총, 총내부, 총외부, 총협, 총송, 외부 / 수급, '예' if 외부 + 1e-9 >= 목표 else '아니오', 수급,
                   f'{기록일.month}/{기록일.day}'))
        print(f'  작성 완료: {아동} (담당 {담당}, {수급}시간형)')

    if 사유없음:
        점검.append(('(전체)', None, f'"계획시간 축소 사유"가 비어 있어요 ({len(사유없음)}명: {", ".join(사유없음[:6])}{" 등" if len(사유없음) > 6 else ""}). 결제내역으로는 알 수 없어서 직접 적어 주세요'))
    for 아동, 글들 in 시간다름.items():
        점검.append((아동, None, f'평소(보통 날 이용시간)와 다른 날이 있어요: {", ".join(글들[:8])}{" 등" if len(글들) > 8 else ""}. 기록지에는 줄어든 시간대로 쓰고, 시간이 줄어든 날의 사유는 설정의 늦은등원_사유/이른하원_사유 글을 썼어요. 실제 사유와 다르면 고쳐 주세요(추가이용 사유는 설정의 추가이용_사유 글을 썼어요)'))
    점검.append(('(전체)', None, '송영서비스가 있는 날은 8월 예시처럼 활동 맨 앞에 "송영서비스" 1시간을 넣고, 그 다음부터 시간표 활동을 했어요'))

    for ws in list(전체.worksheets):
        if not (ws.title.endswith(시트이름_구분 + '계획서') or ws.title.endswith(시트이름_구분 + '기록지')):
            전체.remove(ws)
    전체파일 = os.path.join(출력, f'주간활동_{월}월_계획서_기록지.xlsx')
    전체.save(전체파일)
    print(f'\n  → {폴더이름}/{os.path.basename(전체파일)}  (이용자 {len(대상)}명, 시트 {len(전체.worksheets)}개)')

    wb = Workbook()
    s1 = wb.active
    s1.title = '요약'
    s1.append(['이용자', '담당 제공인력', '서비스일수', '인정결제 일수', '총 계획시간', '내부(시간)', '외부(시간)', '협력(시간)', '송영(시간)',
               '외부+협력 ÷ 수급시간', '30% 이상', '수급시간', '기록지 작성일자'])
    for r in 요약:
        s1.append(list(r))
    for row in s1.iter_rows(min_row=2, min_col=10, max_col=10):
        for c in row:
            c.number_format = '0.0%'
    s2 = wb.create_sheet('확인할 점')
    s2.append(['이용자', '날짜', '내용'])
    본 = set()
    for 아동, d, msg in 점검:
        키 = (아동, d, msg)
        if 키 in 본:
            continue
        본.add(키)
        s2.append([아동, f'{d.month}/{d.day}({요일글자[d.weekday()]})' if d else '(전체)', msg])
    for ws in (s1, s2):
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill('solid', fgColor='DDEBF7')
            c.alignment = Alignment(wrap_text=True, vertical='center')
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 16
    s2.column_dimensions['C'].width = 80
    wb.save(os.path.join(출력, '점검표.xlsx'))
    print(f'\n끝! 이용자 {len(대상)}명 · 확인할 점 {len(본)}건')
    return dict(출력=출력, 폴더이름=폴더이름, 월=월, 아동수=len(대상), 확인할점=len(본), 전체파일=전체파일,
                점검표=os.path.join(출력, '점검표.xlsx'),
                요약=[dict(아동=r[0], 담당=r[1], 서비스일수=r[2], 인정일수=r[3], 총시간=r[4], 외부시간=r[6] + r[7], 통과=(r[10] == '예')) for r in 요약])


if __name__ == '__main__':
    try:
        main()
    except SystemExit as e:
        if e.code not in (None, 0):
            print('\n※', e.code)
    except Exception:
        import traceback
        traceback.print_exc()
        print('\n※ 오류가 났어요. 이 화면을 캡처해서 Claude에게 보여주세요.')
