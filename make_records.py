# -*- coding: utf-8 -*-
"""
방과후활동서비스 월별 활동계획 및 기록지 자동 작성 프로그램  (현양식 버전)
=====================================================================
[하는 일]
  폴더 안의 엑셀 3개(결제내역, 월별 시간표, 현양식)를 자동으로 찾아서 읽고,
  내 담당 아동마다 '계획서 시트 + 기록지 시트'가 들어 있는 엑셀 1개를 '결과' 폴더에 만듭니다.

[지켜지는 규칙]
  - 결제 시간의 .5 는 30분입니다. (2.5 = 2시간 30분)
  - 같은 선생님의 아동들은 같은 날짜·같은 시간에 같은 활동을 합니다. (선생님 몸은 하나)
  - 자체 외부 + 협력기관 시간이 수급시간(66시간)의 30%를 넘고, 총 활동시간의 30%도 넘게 맞춥니다.
    모자라면 자율활동/일상생활훈련/태블릿학습게임/개별학습 칸을 '산책'으로 바꿉니다. (주활동은 건드리지 않음)
  - 요리·미술·과학·국어·체육 같은 주활동을 먼저 배치합니다.
  - 산책은 중랑천, 묵동어린이공원, 화랑대 철도공원을 날짜마다 번갈아 갑니다.
  - 금요일 난타는 외부 2시간, '난타(총 8명)'으로 적습니다.
  - 토요일 외부활동은 10:00~15:00, 그 밖의 토요일 시간은 자율활동/태블릿학습게임입니다.
    토요일이 '실내요리'인 날은  10~11 마트 장보기, 11~12 요리, 12~13 점심식사, 13~14 일상생활훈련(뒷마무리).
  - 정주희·이지윤·문소현·정은지·이동욱 선생님은 평일 15:00부터 근무한 것으로 적습니다. (시간 수는 결제 그대로)
  - 계획서 시트에는 '활동기록'을 비워두고, 기록지 시트에만 채웁니다.
  - 기록지 시트의 개별 특이사항에는 소급결제 사유를 적습니다.
  - 인정결제는 활동 없이 '인정결제 00시간'만 적고, 그 줄 바로 아래 병합된 칸에 요구사항을 적습니다.
  - 작성일자: 계획서 = 전달 마지막 날(일요일·공휴일 제외), 기록지 = 인정결제된 날.
  - 채워 넣는 글씨는 모두 검은색입니다.

[사용법]  같은 폴더에 이 파일과 엑셀 3개를 두고 '실행하기.bat' 을 더블클릭하세요.
"""
import os
import re
import sys
import math
import random
import hashlib
import calendar
import datetime as dt
from copy import copy
from collections import Counter, defaultdict

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.styles.cell_style import StyleArray
from openpyxl.utils import get_column_letter


_고친파일들 = {}


def 엑셀열기(경로, **kw):
    """엑셀을 엽니다. 한글(한컴오피스)에서 저장한 파일처럼 openpyxl 이 못 여는 파일은 자동으로 고쳐서 엽니다."""
    try:
        return load_workbook(경로, **kw)
    except Exception:
        import zipfile
        import tempfile
        고침 = _고친파일들.get(경로)
        if not 고침:
            고침 = os.path.join(tempfile.mkdtemp(), os.path.basename(경로))
            with zipfile.ZipFile(경로) as zin, zipfile.ZipFile(고침, 'w', zipfile.ZIP_DEFLATED) as zout:
                for 항목 in zin.infolist():
                    자료 = zin.read(항목.filename)
                    if 항목.filename.endswith('.xml'):
                        글 = 자료.decode('utf-8', errors='ignore')
                        if 'AlternateContent' in 글:        # 한컴 전용 확장 태그는 대체 내용(Fallback)만 남깁니다
                            글 = re.sub(r'<mc:AlternateContent[^>]*>.*?<mc:Fallback>(.*?)</mc:Fallback>\s*</mc:AlternateContent>',
                                       r'\1', 글, flags=re.S)
                            자료 = 글.encode('utf-8')
                    zout.writestr(항목, 자료)
            _고친파일들[경로] = 고침
        return load_workbook(고침, **kw)

# =====================================================================
# [설정]  ← 바꿀 일이 생기면 여기만 고치면 됩니다
# =====================================================================
내_이름 = '황은미'                        # 결제내역의 '제공인력명'에 적힌 내 이름
연도 = None                               # None 이면 자동 (결제내역에는 연도가 없어서 오늘 날짜로 판단)
결제내역_파일 = ''                         # 비워두면 폴더에서 자동으로 찾습니다
시간표_파일 = ''                           # 비워두면 자동 (이름에 '9월'처럼 이번 달이 들어 있는 파일을 먼저 고릅니다)
양식_파일 = ''                             # 비워두면 자동 (첫 줄에 '활동계획서'가 있는 파일)
기관명 = '재활상담복지센터'
저장_위치 = '바탕화면'                      # '바탕화면' 이면 바탕화면에 폴더를 만듭니다 (프로그램 폴더 옆에 만들고 싶으면 '프로그램폴더')
지점이름 = ''                              # 비우면 폴더 이름이 '방과후기록지_9월', 적으면 '방과후기록지_9월_지점이름' (화면 프로그램에서 입력해도 됩니다)
화면에서_실행 = False                      # 화면(app.py)에서 실행할 때 True 로 바뀝니다
출력_폴더 = '방과후기록지'                  # 바탕화면에 만들어질 폴더 이름. 뒤에 '_9월' 처럼 달이 붙습니다

모든_아동 = True                          # True = 결제내역의 모든 아동(모든 선생님 담당)을 만듭니다. False = 내_이름 선생님 담당 아동만
아동_직접지정 = []                        # 예: ['김수호', '이하나']  비워두면 '내 담당 아동'을 자동으로 찾습니다
대체인력_문구 = '담당 제공인력 휴가'   # 담당이 아닌 선생님이 들어간 날 '개별 특이사항'에 적을 말
계획시간_축소사유 = '치료 센터'            # 총 계획시간이 수급시간보다 적을 때 적을 말
축소사유_지정 = {}                    # 아동마다 다른 축소 사유. 예: {'김하나': '- 치료 일정\n- 학원 일정'}  (없는 아동은 위의 계획시간_축소사유)

기본_시작시각 = 14                         # 평일 활동 시작 (14 = 14:00)
늦은출근_선생님 = {'정주희': 15, '이지윤': 15, '문소현': 15, '정은지': 15, '이동욱': 15}   # 이 선생님들은 15:00 시작
토요일_시작시각 = 9
토요일_외부시작 = 10                       # 토요일 외부활동 10:00 ~ 15:00
토요일_외부끝 = 15
산책_장소들 = ['중랑천', '묵동어린이공원', '화랑대 철도공원']   # 날짜마다 번갈아 갑니다
난타_시간 = 2
난타_인원 = 8
시트이름_구분 = ' '                   # 시트 이름 '아동 계획서' 의 가운데 글자. 광진점은 '_' (예: 구교현_계획서)
외부_최소비율 = 0.30                       # 자체 외부 + 협력기관 ≥ 이 비율을 '넘어야' 합니다
추가_공휴일 = []                           # 공휴일은 자동입니다. (정부가 임시공휴일을 새로 정했을 때만 필요, 보통 비워 두세요)

# 시간표에 적힌 활동 이름에 이 글자가 들어 있으면 자동으로 이렇게 처리합니다 (활동 이름이 바뀌어도 글자만 들어 있으면 따라갑니다)
외부활동_글자 = ['지역사회', '난타', '산책', '외부', '나들이', '견학', '현장학습']   # 센터 밖에서 하는 활동
활동_길이 = {'난타': 난타_시간}                # 기본은 1시간, 다른 길이는 여기에 (이름에 들어 있는 글자: 시간)
활동_뒤 = {'난타': f'(총 {난타_인원}명)'}        # 활동 글 뒤에 붙일 말

# ---- 지점별 옵션 (기본값은 모두 '꺼짐' = 중랑점 방식. 노원점 등은 지점설정.txt 에서 켭니다) ----
결제시각으로_시작 = False          # True = 결제내역의 '시작 시각'(반올림)부터 활동을 시작합니다
시각_반올림_분 = 30                # 14:45~15:14 → 15:00, 15:15~15:44 → 15:30
시각_올림 = False                 # True = 결제 시작 시각을 올림해서 써요 (16:14 → 16:30). 양천 손기록지가 이렇게 적혀 있어요
평일_가장이른시각 = 0              # 예: 15 = 평일 활동은 절대 15시 전에 시작하지 않음
시각_표기 = '콜론'                 # '콜론' = 15:30~16:00 / '시' = 15.5시~16시 / '시분' = 15시 30분~16시
활동_구분 = '\n'                   # ' / ' 로 바꾸면 활동을 가로 한 줄로 씁니다
상세로_대체 = False                # True = 시간표 세부내용이 있으면 활동 이름 대신 세부내용만 써요 (양천점)
속_줄바꿈_펴기 = False             # True = 시간표 한 칸 안의 줄바꿈을 띄어쓰기로 바꿔요 (양천점)
빈_요구사항_줄_빼기 = False         # True = 요구사항이 비어 있으면 '이용자 요구사항)' 빈 줄을 만들지 않아요 (양천점)
기록지_역순 = []                   # 예: ['유효린', '허서현'] 이 아이들은 기록지만 날짜를 말일→1일 순서로 써요 (계획서는 그대로, 양천점)
연장_끝까지 = False                # True = 연장 활동을 하루 끝(결제 끝 시각)까지 이어서 써요 (양천점: 18시 이후는 일상생활)
시간대_줄표시 = False             # True = 시간표 첫 칸에 '15 - 16시' 처럼 시간대가 적힌 줄을 주활동 이름 줄로 읽어요 (양천점)
협력_글자 = []                     # 활동 글에 이 글자가 있으면 '협력' 시간으로 셉니다. 예: ['(협)']
상세만_쓰는_글자 = []              # 시간표 세부내용에 이 글자가 있으면 활동 이름 없이 세부내용만 씁니다
붙은_같은_활동_합쳐쓰기 = True       # True = 17:00~18:00 산책 / 18:00~19:00 산책 처럼 붙어 있는 같은 활동은 17:00~19:00 산책 한 줄로 적기
같은_주활동_하루_한번 = True        # True = 시간표에 같은 주활동이 하루에 두 번 적혀 있어도 한 번만 넣기
같은_활동_합치기 = False           # True = 붙어 있는 같은 활동은 한 칸으로 (외부활동 A+B 등)
주활동_먼저 = False                # True = 시간표 주활동을 모두 먼저, 보조활동은 남는 시간에
산책_가능한_주활동 = []            # 이 글자가 이름에 있는 시간표 활동도 30%가 모자라면 산책으로 바꿔요. 예: ['일상생활']
자율활동_글 = '자율활동 트램펄린'
자율활동_지정 = {}                  # 아이마다 다른 자율활동. 예: {'김하나': ['자율활동(캐치볼)', '자율활동(영단어)'], '이둘': '자율활동(한강라면)'}  (없는 아이는 위의 자율활동_글)
개별학습_글 = '개별학습 선긋기'
사업_글자 = ''                     # 결제내역에 사업이 섞여 있을 때 이 글자가 든 사업만 써요. 예: 마포 주간 '발달장애인 주간' (최중증 빼기)
대상_서비스 = '방과후활동서비스'     # 결제내역 '서비스유형'에서 이 서비스를 한 아동만 만들어요. 주간이면 '주간활동서비스' (마포점 주간)
하루_틀_지정 = {}                  # 요일별 하루 순서를 직접 정해요 (마포 주간). 예: {0: [('#1',1), ('산책',1,True), ('자립생활활동지도',1), ('#3',1), ('일상생활 훈련 및 하원',9)]}  '#N' = 시간표의 N번째 묶음
시간표_시트 = ''                   # 시간표 엑셀에 시트가 여러 개일 때 쓸 시트 이름. 예: '방과후시간표' (비우면 '시간표' 시트, 없으면 첫 시트)
토요일_통째외부 = False            # True = 토요일 결제 시간 전체를 외부활동 한 칸으로
토요일_이름 = '외부활동'
토요일_장소들 = {}                   # 예: {'9/5': '고양 낙농치즈체험관'} 시간표 대신 날짜별 토요일 장소를 직접 적어요
모든_운영일_행 = False             # True = 결제가 없는 날도 (일요일·공휴일 빼고) 날짜 줄을 만듭니다
인정결제만_있는_날_줄_빼기 = False   # True = 인정결제만 있고 활동이 없는 날은 날짜 줄을 만들지 않아요 (인정결제 시간은 요구사항 줄에 적을 때 쓰세요)
오른쪽_0_쓰기 = False              # True = 내부/외부/협력/송영 칸에 0도 씁니다
안내문_지울_글자 = []              # 양식 아래 안내문 중 이 글자가 들어 있는 줄은 지웁니다. 예: ['210㎜']
요구사항_맨아래 = False            # True = 요구사항 줄을 날짜 줄들 맨 아래에 한 줄로 ('이용자 요구사항) ...')
요구사항_머리말 = '이용자 요구사항) '
인정결제_요구사항에 = False        # True = 인정결제 시간을 요구사항 줄 끝에 적고 특이사항에는 적지 않아요
기록일_방식 = '인정결제'           # '인정결제' = 인정결제된 날 / '말일' = 그 달 마지막 근무일
소급_문구 = {}                     # 결제 사유 → 특이사항에 쓸 말. 예: {'대상자 과실(카드 미소지 등)': '카드 미소지로 소급결제'}
대체_기록지에만 = False            # True = 대체인력 안내를 기록지에만 적고 계획서에는 적지 않아요
일부만_하는_활동 = {}              # 시간표에 적힌 활동 중 '몇 명만' 하는 활동. 예: {'카엘': 4} = 그 활동이 있는 날 랜덤 4명만 하고 나머지는 대신 활동
일부만_대신_활동 = ''              # 나머지 아이가 대신 하는 활동 글 (비우면 자율활동)
수급시간_바꾸기 = {}               # 개인예산제 등. 예: {'김신우': 52.5}
수급시간_칸_채우기 = False         # True = 양식의 '수급시간' 칸을 ■ 월 66시간 으로 채웁니다
송영_글에_쓰기 = True               # False = 송영서비스를 활동 글에는 쓰지 않고 송영 칸(O)에만 세요 (마포점)
송영_따로_처리 = False             # True = 결제내역의 '송영서비스'를 활동 시간과 따로 세고 송영 칸(O)에 씁니다
공통_시작_고정 = False             # True = 같은 반 시간표는 항상 기본 시작 시각부터 (아동은 자기 결제 시각 이후 부분만 씁니다)
하루_틀_시간 = 0                   # 예: 4 = 하루 기본 틀을 정확히 4시간(15~19시)으로 짜고 그 뒤는 '연장'. 0이면 사용 안 함
보조_맨앞_요일 = []                # 자율활동을 하루 맨 앞에 두는 요일 (0=월 ... 5=토). 예: [0, 1]
협력_먼저 = False                  # True = 협력기관 활동을 주활동 중 가장 먼저
연장_활동 = {}                     # 하루 기본 틀 뒤 첫 시간에 하는 활동. 예: {0: '요리활동-스파게티', 1: '요리활동-야채볶음밥'}
활동_연결 = '('                    # '(' → 미술활동(가전제품색칠) / '-' → 미술활동-가전제품색칠
토요일_글_모양 = ''                # 예: '{이름} ({장소})' → 외부활동 (꼬꼬체험관)
날짜_모양 = ''                     # '붙임' → 2026.8.31. (월)
날짜칸_날짜형 = False               # True = 맨 아래 작성일을 글자가 아닌 진짜 날짜 칸으로 써요 (마포점)
인정결제_문구 = '인정결제 {n}시간'   # 인정결제 줄에 적을 말. 마포점: '인정결제 : {n} 시간'
인정결제_덧붙임 = ''               # 인정결제 줄 아래에 한 줄 더 적을 말. 마포점: '이용자 요구사항 : 원함'
인정결제_표기 = True               # False = 인정결제 시간을 어디에도 적지 않아요
요구사항_자동 = True               # False = 요구사항 문구를 프로그램이 쓰지 않고 칸만 비워 둬요 (중랑 말고 다른 지점은 직접 적어요)
요구사항_특이사항 = False          # True = 요구사항을 기록지의 마지막 날짜 줄 '개별 특이사항'에 '요구사항: ...' 로 써요
양천식_글 = False                   # True = 대제목(소제목) 글 모양을 양천 손기록지처럼: 난타(풍선…) / 요리활동(퀘사디아) / 안쪽 괄호는 ' - '
소제목_빼기 = []                    # 이 글자가 든 활동은 세부(소제목)를 쓰지 않고 이름만 적어요 (양천: 스마트교육센터)
빈내용_알림 = True                 # False = 시간표 칸에 세부내용이 없어도 점검표에 알리지 않아요 (핑거푸드처럼 원래 이름만 쓰는 활동이 많을 때)
담당_지정 = {}                     # 담당 선생님을 직접 정할 때. 예: {'김서진': '손은주'}  (비우면 결제를 가장 많이 한 선생님)

# 시간표 글자를 기록지에 맞게 고치는 목록 (시간표 글자 → 기록지에 쓸 글자)
글자_바꾸기 = {
    '마트장보기': '마트 장보기',
    '클레이아트': '클레이 아트',
}
이름_바꾸기 = {'지역사회': '지역사회활동'}

# 일요일 외에 쉬는 공휴일 (대체공휴일 포함)
공휴일_목록 = """
2025-01-01 2025-01-27 2025-01-28 2025-01-29 2025-01-30 2025-03-01 2025-03-03 2025-05-05 2025-05-06
2025-06-03 2025-06-06 2025-08-15 2025-10-03 2025-10-05 2025-10-06 2025-10-07 2025-10-08 2025-10-09 2025-12-25
2026-01-01 2026-02-16 2026-02-17 2026-02-18 2026-03-01 2026-03-02 2026-05-05 2026-05-24 2026-05-25
2026-06-03 2026-06-06 2026-08-15 2026-08-17 2026-09-24 2026-09-25 2026-09-26 2026-10-03 2026-10-05
2026-10-09 2026-12-25
2027-01-01 2027-02-05 2027-02-06 2027-02-07 2027-02-08 2027-03-01 2027-05-05 2027-05-13 2027-06-06
2027-08-15 2027-08-16 2027-09-14 2027-09-15 2027-09-16 2027-10-03 2027-10-04 2027-10-09 2027-10-11
2027-12-25 2027-12-27
"""

# =====================================================================
# 하루 활동 틀
#   주활동(요리·미술·국어 ...)은 시간표에서 요일마다 자동으로 읽어 옵니다.
#   (요일이 바뀌거나 활동이 바뀌어도 시간표만 새로 넣으면 됩니다)
#   시간표에 없는 보조활동(자율활동/개별학습/산책)만 여기서 정합니다.
# =====================================================================
바꿀수있는_종류 = ('자율', '일상', '태블릿', '개별')       # 30%가 모자랄 때 '산책'으로 바꿀 수 있는 칸
바꿀_차례 = ['자율', '일상', '태블릿', '개별']


def 보조(글, 외부, 종류, 길이=1):
    return dict(kind='보조', text=글, ext=외부, tag=종류, dur=길이, conv=(종류 in 바꿀수있는_종류 and not 외부))


def _고르기(글, 날짜):
    if isinstance(글, (list, tuple)):
        return 글[(날짜.toordinal() if 날짜 else 0) % len(글)]
    return 글


def 자율(날짜=None):
    return 보조(_고르기(자율활동_글, 날짜), False, '자율')


def 개별(날짜=None):
    return 보조(_고르기(개별학습_글, 날짜), False, '개별')


def 산책():
    return 보조('산책', True, '산책')


# 시간표를 읽지 못했을 때만 쓰는 예비 이름
기본_주활동 = {0: ['요리활동', '종이접기'], 1: ['미술활동', '과학활동'], 2: ['체육활동', '소근육활동', '음악활동'],
            3: ['오감활동', '국어활동'], 4: ['공간지각활동', '지역사회활동', '난타']}


덧붙일_활동 = [('자율활동 빈백', '자율'), ('태블릿학습게임', '태블릿'), ('일상생활훈련', '일상')]

요일글자 = ['월', '화', '수', '목', '금', '토', '일']
요일이름 = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일']
요일번호 = {'월': 0, '화': 1, '수': 2, '목': 3, '금': 4, '토': 5}


# =====================================================================
# 도우미 함수
# =====================================================================
def 기준폴더들():
    if getattr(sys, 'frozen', False):
        기준 = os.path.dirname(sys.executable)
    elif '__file__' in globals():
        기준 = os.path.dirname(os.path.abspath(__file__))
    else:
        기준 = os.getcwd()
    return [기준] if 기준 == os.getcwd() else [기준, os.getcwd()]


def 엑셀목록():
    본 = set()
    for 폴더 in 기준폴더들():
        for 이름 in os.listdir(폴더):
            경로 = os.path.join(폴더, 이름)
            if (이름.lower().endswith('.xlsx') and not 이름.startswith('~$') and not 이름.endswith('_기록지.xlsx')
                    and 이름 != '점검표.xlsx' and 경로 not in 본):
                본.add(경로)
                yield 경로


def 종류_알아내기(경로):
    try:
        wb = 엑셀열기(경로, read_only=True, data_only=True)
    except Exception:
        return None
    try:
        첫줄 = next(wb.worksheets[0].iter_rows(min_row=1, max_row=1, values_only=True), ())
        글 = ' '.join(str(v) for v in 첫줄 if v)
        if '활동계획서' not in 글:                    # 양식 위쪽에 빈 줄이 더 있는 지점도 있어서 위 8줄을 봅니다
            글 = ' '.join(str(v) for 줄 in wb.worksheets[0].iter_rows(min_row=1, max_row=8, values_only=True) for v in 줄 if v)
        if '활동계획서' in 글 and '활동기록지' in 글:
            return '양식'
        if '대상자명' in 글 and '제공인력명' in 글:
            return '결제'
        if '요구사항' in wb.sheetnames or '시간표' in wb.sheetnames:
            return '시간표'
        for 줄 in wb.worksheets[0].iter_rows(min_row=1, max_row=3, values_only=True):    # 시트 이름이 달라도 '월 화 수 목 금 토' 줄이 있으면 시간표
            if len({str(v).strip() for v in 줄 if v} & {'월', '화', '수', '목', '금', '토'}) >= 4:
                return '시간표'
    except Exception:
        return None
    finally:
        wb.close()
    return None


def 파일_고르기(설정이름, 종류, 월=None):
    if 설정이름:
        for 폴더 in 기준폴더들():
            경로 = os.path.join(폴더, 설정이름)
            if os.path.exists(경로):
                return 경로
        return None
    후보 = [p for p in 엑셀목록() if 종류_알아내기(p) == 종류]
    if not 후보:
        return None
    후보.sort(key=os.path.getmtime, reverse=True)
    if 월 is not None:
        이번달 = [p for p in 후보 if f'{월}월' in os.path.basename(p)]
        if 이번달:
            return 이번달[0]
    return 후보[0]


def 시간글자(h):
    return f'{h:g}'


def 숫자(h):
    return int(h) if abs(h - round(h)) < 1e-9 else h


def 시각(t):
    if 시각_표기 == '시':
        return f'{숫자(t)}시'
    h = int(t + 1e-9)
    m = 30 if (t - h) >= 0.5 - 1e-9 else 0
    if 시각_표기 == '시분':
        return f'{h}시 30분' if m else f'{h}시'
    return f'{h:02d}:{m:02d}' if 시각_표기 == '콜론0' else f'{h}:{m:02d}'


def 글자폭(s):
    return sum(2 if ord(ch) >= 0x1100 else 1 for ch in s)


def 줄수(text, 칸폭):
    if not text:
        return 1
    return sum(max(1, math.ceil(글자폭(ln) / 칸폭)) for ln in str(text).split('\n'))


def 월요일(d):
    return d - dt.timedelta(days=d.weekday())


def 한줄(날짜):
    return f'{날짜.month}/{날짜.day}({요일글자[날짜.weekday()]})'


def 날짜글(d):
    if 날짜_모양 == '붙임':
        return f'{d.year}.{d.month}.{d.day}. ({요일글자[d.weekday()]})'
    return f'{d.year}. {d.month}. {d.day}. ({요일글자[d.weekday()]})'


_휴일집합 = None
_자동공휴일 = False


def 공휴일들():
    """공휴일을 자동으로 알아냅니다. (holidays 도구가 있으면 그것을 쓰고, 없으면 코드에 넣어 둔 2025~2027년 목록을 씁니다)"""
    global _휴일집합, _자동공휴일
    if _휴일집합 is None:
        s = set(공휴일_목록.split())
        s.update(추가_공휴일)
        try:
            import holidays
            for y in range(2024, 2036):
                try:
                    h = holidays.country_holidays('KR', years=y)
                except AttributeError:
                    h = holidays.KR(years=y)
                s.update(d.isoformat() for d in h.keys())
            _자동공휴일 = True
        except Exception:
            _자동공휴일 = False
        _휴일집합 = s
    return _휴일집합


def 쉬는날(d):
    """일요일이거나 공휴일이면 True (토요일은 센터가 운영하므로 쉬는 날이 아님)"""
    return d.weekday() == 6 or d.isoformat() in 공휴일들()


def 전달_마지막_근무일(년, 월):
    d = dt.date(년, 월, 1) - dt.timedelta(days=1)
    while 쉬는날(d):
        d -= dt.timedelta(days=1)
    return d


def 이번달_마지막_근무일(년, 월):
    d = dt.date(년, 월, calendar.monthrange(년, 월)[1])
    while 쉬는날(d):
        d -= dt.timedelta(days=1)
    return d


def 검정(셀):
    f = 셀.font
    셀.font = Font(name=f.name, size=f.sz, bold=f.b, italic=f.i, underline=f.u,
                  strike=f.strike, vertAlign=f.vertAlign, color='FF000000')


# =====================================================================
# 1. 결제내역 읽기
# =====================================================================
결제_건너뜀 = []


def 결제내역_새형식_바꾸기(rows, 머리, 위치):
    """'승인일시 / 총시간(00300 = 3시간) / 소급사유 / 결제구분(승인취소·매입취소)' 모양의 결제내역을
    기존 모양('시작시간(월/일/분)', '서비스 제공시간 합계', '일괄결제(사유)')으로 바꿉니다.
    승인취소·매입취소 줄(취소된 결제)은 뺍니다."""
    def 칸(r, 이름):
        i = 위치.get(이름)
        v = r[i] if i is not None and i < len(r) else None
        return '' if v is None else str(v).strip()

    def 시간(v):
        t = re.sub(r'\D', '', v).zfill(5)
        return int(t[:-2]) + int(t[-2:]) / 60

    자료 = []
    for r in rows[1:]:
        if not r or not 칸(r, '대상자명'):
            continue
        자료.append(dict(아동=칸(r, '대상자명'), 유형=칸(r, '서비스유형'), 사업=칸(r, '사업유형'), 인력=칸(r, '제공인력명'),
                         구분=칸(r, '결제구분'), 시작=칸(r, '시작시간'), 종료=칸(r, '종료시간'), 시간=시간(칸(r, '총시간')),
                         사유=칸(r, '소급사유'), 승인=칸(r, '승인일시')))
    살림 = [x['구분'] not in ('승인취소', '매입취소') for x in 자료]      # 취소된 결제 줄은 그대로 뺍니다
    새머리 = ['대상자명', '제공인력명', '사업유형', '서비스유형', '시작시간(월/일/분)', '서비스 제공시간 합계', '일괄결제(사유)', '반납구분']
    새 = [새머리]
    for ok, x in zip(살림, 자료):
        if ok:
            새.append([x['아동'], x['인력'], x['사업'], x['유형'], x['시작'], f"{x['시간']:.2f}", x['사유'], ''])
    return 새


def 결제내역_읽기(경로):
    wb = 엑셀열기(경로, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    머리 = [str(h).strip() if h is not None else '' for h in rows[0]]
    위치 = {h: i for i, h in enumerate(머리)}
    if '서비스 제공시간 합계' not in 위치 and '총시간' in 위치 and '시작시간' in 위치:
        rows = 결제내역_새형식_바꾸기(rows, 머리, 위치)       # 승인일시·총시간 모양의 새 결제내역(양천 등)
        머리 = [str(h).strip() if h is not None else '' for h in rows[0]]
        위치 = {h: i for i, h in enumerate(머리)}
    for 이름 in ['대상자명', '제공인력명', '시작시간(월/일/분)', '서비스 제공시간 합계']:
        if 이름 not in 위치:
            raise SystemExit(f'결제내역에서 "{이름}" 칸을 찾지 못했어요. 첫 줄 제목을 확인해 주세요.')

    def 값(r, 이름):
        i = 위치.get(이름)
        v = r[i] if i is not None and i < len(r) else None
        return '' if v is None else str(v).strip()

    # 파일에 따라 '제목 줄'이 첫 줄이 아닐 수 있어 위쪽 몇 줄 안에서 제목 줄을 찾습니다 (위에서 못 찾았을 때)
    결과 = []
    유형칸있음 = '서비스유형' in 위치
    for r in rows[1:]:
        if not r or not 값(r, '대상자명'):
            continue
        원본 = r[위치['시작시간(월/일/분)']]
        시각값 = None
        if isinstance(원본, (dt.datetime, dt.date)):
            월, 일 = 원본.month, 원본.day
            if isinstance(원본, dt.datetime):
                시각값 = 원본.hour + 원본.minute / 60
        else:
            m = re.match(r'\s*(\d{1,2})[-/.](\d{1,2})', str(원본))
            if not m:
                결제_건너뜀.append((값(r, '대상자명'), f'날짜 모양이 달라 건너뛴 줄: {값(r, "결제구분")} {원본} ({값(r, "서비스 제공시간 합계")}시간) - 다른 달 예외결제일 수 있어요'))
                continue
            월, 일 = int(m.group(1)), int(m.group(2))
            mt = re.search(r'(\d{1,2}):(\d{2})', str(원본))
            if mt:
                시각값 = int(mt.group(1)) + int(mt.group(2)) / 60
        try:
            시간 = float(값(r, '서비스 제공시간 합계'))
        except ValueError:
            결제_건너뜀.append((값(r, '대상자명'), f'시간이 숫자가 아니라 건너뛴 줄: {원본} "{값(r, "서비스 제공시간 합계")}"'))
            continue
        결과.append(dict(아동=값(r, '대상자명'), 인력=값(r, '제공인력명'),
                         유형=값(r, '서비스유형') if 유형칸있음 else '방과후활동서비스',
                         월=월, 일=일, 시간=시간, 사유=값(r, '일괄결제(사유)'), 반납=값(r, '반납구분'), 시각=시각값, 사업=값(r, '사업유형')))
    return 결과


# =====================================================================
# 2. 시간표 / 요구사항 읽기
# =====================================================================
class 시간표:
    """시간표 엑셀을 읽습니다. 요일 칸(월~토), 주차 줄, 주활동 이름 줄을 모양으로 알아내므로
    요일이 바뀌거나 활동이 바뀌거나 활동이 늘고 줄어도 자동으로 따라갑니다."""

    def __init__(self, 경로, 연, 월):
        self.blocks = []
        self.ok = False
        self.mode = 'date'
        self.연, self.월 = 연, 월
        마지막 = calendar.monthrange(연, 월)[1]
        self.mondays = sorted({월요일(dt.date(연, 월, k)) for k in range(1, 마지막 + 1)})
        if not 경로:
            return
        wb = 엑셀열기(경로, data_only=True)
        ws = (wb[시간표_시트] if 시간표_시트 and 시간표_시트 in wb.sheetnames else
              wb['시간표'] if '시간표' in wb.sheetnames else wb[wb.sheetnames[0]])
        grid = {}
        for row in ws.iter_rows():
            for c in row:
                grid[(c.row, c.column)] = c.value

        열 = {}
        머리줄 = 1
        for r in range(1, 8):                              # 월 화 수 목 금 토 가 적힌 줄 찾기
            찾음 = {요일번호[str(grid.get((r, c))).strip()]: c
                   for c in range(1, 14) if str(grid.get((r, c))).strip() in 요일번호}
            if len(찾음) >= 4:
                열, 머리줄 = 찾음, r
                break
        if not 열:
            return

        띠행 = set()                                       # 요일 칸 전체가 한 칸으로 합쳐진 줄 = 보조활동 띠(자율활동&산책 ...)
        for mr in ws.merged_cells.ranges:
            덮음 = [c for c in 열.values() if mr.min_col <= c <= mr.max_col]
            if mr.min_row == mr.max_row and len(덮음) >= max(2, len(열) - 1):
                띠행.add(mr.min_row)
        for mr in list(ws.merged_cells.ranges):            # 병합된 칸은 같은 값으로 채움
            v = ws.cell(mr.min_row, mr.min_col).value
            for r in range(mr.min_row, mr.max_row + 1):
                for cc in range(mr.min_col, mr.max_col + 1):
                    grid[(r, cc)] = v

        self.순서만 = False
        현재 = None
        for r in range(머리줄 + 1, ws.max_row + 1):
            if r in 띠행:
                continue
            a = grid.get((r, 1))
            값들 = {wd: ('' if grid.get((r, c)) is None else str(grid.get((r, c))).strip())
                    for wd, c in 열.items()}
            채움 = [v for v in 값들.values() if v]
            주차 = self._주파싱(a)
            시각줄 = (주차 is None and isinstance(a, str) and 채움 and 시간대_줄표시
                     and re.match(r'\s*\d{1,2}(:\d{2})?\s*[-~]\s*\d{1,2}', a) is not None)
            if 시각줄:                                      # '15 - 16시' 처럼 시간대가 적힌 줄 = 주활동 이름 줄(새 묶음)
                현재 = {'cat': 값들, 'weeks': []}
                self.blocks.append(현재)
                continue
            if 주차 is None and isinstance(a, str) and a.strip() and 현재 is not None and 채움:
                주차 = (0, 0)                                # '1주차' 처럼 날짜가 아닌 이름 → 순서대로 맞춤
                self.순서만 = True
            if 주차:
                if 현재 is not None:
                    현재['weeks'].append((주차, 값들))
            elif 채움 and a in (None, ''):                  # 날짜 없는 줄 = 주활동 이름 줄(새 묶음)
                현재 = {'cat': 값들, 'weeks': []}
                self.blocks.append(현재)
        self.ok = bool(self.blocks)

        if self.ok:
            라벨들 = {self._라벨_월요일(md) for 블 in self.blocks for md, _ in 블['weeks']}
            맞음 = len(라벨들 & set(self.mondays))
            if self.순서만 or 맞음 / max(1, len(self.mondays)) < 0.6:
                self.mode = 'order'

    @staticmethod
    def _주파싱(a):
        if isinstance(a, (dt.datetime, dt.date)):
            return (a.month, a.day)
        if isinstance(a, str):
            m = re.search(r'(\d{1,2})\s*/\s*(\d{1,2})', a)
            if m:
                return (int(m.group(1)), int(m.group(2)))
        return None

    def _라벨_월요일(self, md):
        m, d = md
        기준 = dt.date(self.연, self.월, 15)
        후보 = []
        for y in (self.연 - 1, self.연, self.연 + 1):
            try:
                후보.append(dt.date(y, m, d))
            except ValueError:
                pass
        if not 후보:
            return None
        return 월요일(min(후보, key=lambda x: abs((x - 기준).days)))

    def _주번호(self, 블록, 날짜):
        주들 = self.blocks[블록]['weeks']
        월 = 월요일(날짜)
        if self.mode == 'date':
            for i, (md, _) in enumerate(주들):
                if self._라벨_월요일(md) == 월:
                    return i
            return None
        순서 = self.mondays.index(월) if 월 in self.mondays else 0
        return min(순서, len(주들) - 1)

    def 값(self, 블록, 날짜, 요일, 빈칸_채우기=True):
        """(시간표 윗줄 이름, 그 주의 내용, x표시인지, 다른 주에서 보충했는지)"""
        if not self.ok or 블록 >= len(self.blocks):
            return '', '', False, False
        칸 = self.blocks[블록]
        헤더 = 칸['cat'].get(요일, '')
        주들 = 칸['weeks']
        i = self._주번호(블록, 날짜)
        if not 주들 or i is None:
            return 헤더, '', False, False
        원문 = 주들[i][1].get(요일, '')
        if 원문.lower() in ('x', '×', 'ｘ'):
            return 헤더, '', True, False
        if 원문 or not 빈칸_채우기:
            return 헤더, 원문, False, False
        후보 = list(range(i - 1, -1, -1)) + list(range(i + 1, len(주들)))
        for j in 후보:
            v = 주들[j][1].get(요일, '')
            if v and v.lower() not in ('x', '×', 'ｘ'):
                return 헤더, v, False, True
        return 헤더, '', False, False

    def 주활동_목록(self, 날짜):
        """그 날(요일)의 주활동들. 각 항목: 묶음 번호, 이름, 그 주의 세부내용"""
        wd = 날짜.weekday()
        결과 = []
        for b in range(len(self.blocks)):
            헤더, 내용, x, 보충 = self.값(b, 날짜, wd)
            H = [p.strip() for p in 헤더.split('/') if p.strip()]
            C = [p.strip() for p in 내용.split('/') if p.strip()]
            if C:
                for p, 속 in enumerate(C):
                    결과.append(dict(block=b, 이름=H[p] if p < len(H) else '', 속=속, x=False,
                                   보충=보충 and p == 0, 빈내용=False))
            elif H:
                결과.append(dict(block=b, 이름=H[0], 속='', x=x, 보충=False, 빈내용=not x))
        return 결과

    def 토요일_장소(self, 날짜):
        for 블록 in range(len(self.blocks)):
            _, 내용, x, _ = self.값(블록, 날짜, 5, 빈칸_채우기=False)
            if 내용:
                return 내용
        return ''


def 요구사항_읽기(경로):
    if not 경로:
        return {}
    wb = 엑셀열기(경로, data_only=True)
    if '요구사항' not in wb.sheetnames:
        return {}
    결과 = defaultdict(list)
    for r in wb['요구사항'].iter_rows(min_row=4, values_only=True):
        if r and len(r) >= 4 and r[0] and r[3]:
            결과[str(r[0]).strip()].append((str(r[2] or '').strip(), str(r[3]).strip()))
    return 결과


def 요구사항_글(아동, 이름들, 목록, 산책있음):
    """아동이 다니는 요일마다, 그 요일에 하는 주활동 이름에 맞는 요구사항 문구를 하나씩 고릅니다.
    (활동이 다른 요일로 옮겨 가도 활동 이름으로 찾으니 문구가 어긋나지 않아요)"""
    줄 = []

    def 맞음(이름, 영역, 문구):
        return bool(이름) and ((영역 and (이름 == 영역 or 이름 in 영역 or 영역 in 이름)) or 이름 in 문구)

    def 고르기(후보, 씨):
        후보 = list(dict.fromkeys(후보))
        if 후보:
            n = int(hashlib.md5(f'{아동}|{씨}'.encode('utf-8')).hexdigest()[:8], 16)
            줄.append('- ' + 후보[n % len(후보)])

    모두 = []
    for wd in sorted(이름들):
        후보 = [문구 for 키, 리스트 in 목록.items() if 키 != '공통(산책)'
              for 영역, 문구 in 리스트 if any(맞음(n, 영역, 문구) for n in 이름들[wd])]
        if not 후보 and wd < 6:
            후보 = [문구 for _, 문구 in 목록.get(요일이름[wd], [])]
        모두 += 후보
    if not 모두 and 산책있음:
        모두 = [문구 for _, 문구 in 목록.get('공통(산책)', [])]
    고르기(모두, '요구')           # 요구사항은 1개만 적습니다
    return '이용자(보호자) 요구사항\n' + '\n'.join(줄) if 줄 else '이용자(보호자) 요구사항'



# =====================================================================
# 3. 하루 활동 만들기  (날짜마다 '공통 시간표' 1개를 만들고, 아동은 자기 시간만큼 잘라서 씁니다)
# =====================================================================
def 글자고치기(s):
    for a, b in 글자_바꾸기.items():
        s = s.replace(a, b)
    return s


def _안쪽괄호(속):
    return re.sub(r'\s*\(([^()]*)\)', r' - \1', 속).strip()


def _양천식_글(머리, 속):
    """머리=시간표 큰칸(대제목), 속=주차칸(소제목). 손기록지처럼 대제목(소제목) 모양으로 써요"""
    if not 속 or 속.replace(' ', '') == 머리.replace(' ', ''):
        return 머리 or 속
    if any(k in 속 for k in 상세만_쓰는_글자):
        return 속
    조각 = [x.strip() for x in re.split(r'\s+&\s+', 머리) if x.strip()]
    속조각 = [x.strip() for x in re.split(r'\s+&\s+(?![^()]*\))', 속) if x.strip()]
    # 가) 소제목이 대제목 이름으로 시작하면 (난타 (풍선…) / 스마트교육센터 (AI…) & 요리 (퀘사디아)) 이름 바로 뒤에 괄호
    def 앞(q):
        return re.split(r'\s*\(', q, 1)[0].strip()
    짝 = {}
    for q in 속조각:
        if q in 조각:
            continue                                        # 소제목 끝에 대제목 이름이 또 붙은 조각(& 양천 숲 체험)은 짝짓기에서 빼요
        for i, p_ in enumerate(조각):
            h = 앞(q)
            if h and (p_.startswith(h) or h.startswith(p_)) and i not in 짝:
                짝[i] = q
                break
    if 짝:
        나옴 = []
        for i, p_ in enumerate(조각):
            q = 짝.get(i)
            if not q:
                continue                                    # 이 주에 세부내용이 없는 대제목 조각은 안 써요 (예: 요리활동이 없는 주)
            꼬리 = ''
            m = re.search(r'\(([^()]*(?:\([^()]*\))?[^()]*)\)\s*$', q)
            꼬리 = _안쪽괄호(m.group(1)) if m else ''
            if any(k in p_ for k in 소제목_빼기):
                꼬리 = ''
            나옴.append(f'{p_}({꼬리})' if 꼬리 else p_)
        for q in 속조각:
            if q not in 짝.values():
                나옴.append(re.sub(r'\s+\(', '(', q))
        return ' & '.join(나옴)
    # 나) 아니면 대제목(소제목). 소제목 끝에 대제목 이름이 또 붙어 있으면 괄호 밖으로
    뒤 = ''
    for p_ in 조각:
        m = re.search(r'\s*&\s*' + re.escape(p_) + r'\s*$', 속)
        if m:
            뒤 += ' & ' + p_
            속 = 속[:m.start()].strip()
    속 = _안쪽괄호(속)
    return (f'{머리}({속})' if 속 else 머리) + 뒤


def 활동글(이름, 속):
    머리 = 이름_바꾸기.get(이름, 이름)
    속 = 글자고치기(속)
    if 속_줄바꿈_펴기 and 속:
        속 = re.sub(r'\s*\n\s*', ' ', 속).strip()
    if 양천식_글:
        return _양천식_글(글자고치기(머리), 속)
    if 상세로_대체 and 속:
        return 속
    if 속 and any(k in 속 for k in 상세만_쓰는_글자):
        return 속
    if not 속 or 속.replace(' ', '') == 머리.replace(' ', ''):
        return 머리 or 속
    if not 머리:
        return 속
    if 활동_연결 != '(':
        return f'{머리}{활동_연결}{속}'
    return f'{머리}({속})'


def 사전에서(이름, 속, 사전, 기본):
    for 열쇠, 값 in 사전.items():
        if 열쇠 in (이름 + 속):
            return 값
    return 기본


def 오늘의_주활동(날짜, tt, 주의):
    """그 날 시간표에 적힌 주활동을 자동으로 읽어 옵니다. (요일·활동·개수가 바뀌어도 시간표대로 따라갑니다)"""
    wd = 날짜.weekday()
    if tt.ok:
        항목 = tt.주활동_목록(날짜)
    else:
        항목 = [dict(block=0 if i == 0 else 1, 이름=n, 속='', x=False, 보충=False, 빈내용=False)
              for i, n in enumerate(기본_주활동.get(wd, []))]
    결과 = []
    for it in 항목:
        글 = 활동글(it['이름'], it['속'])
        if tt.ok:
            if it['x']:
                주의.append((날짜, f'시간표에 x 표시라서 "{글}"(으)로만 적었어요. 확인해 주세요'))
            elif it['빈내용'] and 빈내용_알림:
                주의.append((날짜, f'시간표에 "{글}" 내용이 비어 있어 이름만 적었어요'))
            elif it['보충']:
                주의.append((날짜, f'시간표 칸이 비어서 다른 주의 "{글}"를 가져왔어요'))
        기준 = it['이름']                       # 외부 여부는 활동 '이름'으로만 판단 (세부내용의 '체험' 같은 글자에 속지 않도록)
        if 상세로_대체 and it['속']:
            기준 = 글
        글 = 글 + 사전에서(it['이름'], it['속'], 활동_뒤, '')
        협력 = any(k in 글 for k in 협력_글자)
        바꿀 = any(k in 기준 for k in 산책_가능한_주활동) and not 협력
        결과.append(dict(block=it['block'], 이름=이름_바꾸기.get(it['이름'], it['이름']),
                       text=글, coop=협력,
                       ext=any(k in 기준 for k in 외부활동_글자) and not 협력,
                       dur=사전에서(it['이름'], it['속'], 활동_길이, 1),
                       tag='일상' if 바꿀 else '주', conv=바꿀))
    if 같은_주활동_하루_한번:
        본 = set()
        유일 = []
        for m in 결과:
            if m['text'] in 본:
                continue
            본.add(m['text'])
            유일.append(m)
        결과 = 유일
    if tt.ok and wd < 5 and not 결과:
        주의.append((날짜, '이 요일에는 시간표에 주활동이 없어 보조활동으로만 채웠어요'))
    return 결과


def 평일_공통(날짜, 시작, 최소, 최대, tt, 주의):
    """그 날 아동들이 함께 하는 활동 목록. 최소/최대 = 그 날 아동들의 결제 시간 중 가장 짧은/긴 것
    순서: 첫째 주활동 묶음 → 자율활동 → 둘째 주활동 묶음 → 개별학습 → ... → 산책"""
    주활동들 = 오늘의_주활동(날짜, tt, 주의)
    for 글 in 일부만_하는_활동:                 # '마트/카엘댄스' 처럼 한 칸에 나란히 적힌 것 = 둘 중 하나 (몇 명만 뒤쪽을 함)
        for m in [x for x in 주활동들 if 글 in x['text']]:
            짝 = [x for x in 주활동들 if x['block'] == m['block'] and 글 not in x['text']]
            if 짝:
                짝[0]['alt'] = dict(text=m['text'], coop=m['coop'], ext=m['ext'])
                주활동들.remove(m)
    블록들 = sorted({m['block'] for m in 주활동들})
    틀 = []
    if 하루_틀_시간:                      # 하루 기본 틀을 정확히 N시간으로: 주활동 + 모자라는 만큼 자율활동/개별학습
        주 = [dict(kind='주', text=m['text'], ext=m['ext'], dur=m['dur'], tag=m['tag'], conv=m['conv'], coop=m['coop'], alt=m.get('alt'))
             for m in sorted(주활동들, key=lambda m: m['block'])]
        if 협력_먼저:
            주 = [m for m in 주 if m['coop']] + [m for m in 주 if not m['coop']]
        합 = 0
        잘린 = []
        for m in 주:
            if 합 >= 하루_틀_시간 - 1e-9:
                break
            m['dur'] = min(m['dur'], 하루_틀_시간 - 합)
            합 += m['dur']
            잘린.append(m)
        부족 = max(0, round(하루_틀_시간 - 합))
        채움 = [(자율(날짜), 개별(날짜))[i % 2] for i in range(부족)]
        if 채움 and 날짜.weekday() in 보조_맨앞_요일:
            틀 = [채움[0]] + 잘린 + 채움[1:]
        else:
            틀 = 잘린 + 채움
        if 연장_활동.get(날짜.weekday()):
            틀.append(dict(kind='보조', text=연장_활동[날짜.weekday()], ext=False, dur=(9 if 연장_끝까지 else 1), tag='연장', conv=False, coop=False))
        최소 = 0
        블록들 = []
    for i, b in enumerate(블록들):
        틀 += [dict(kind='주', text=m['text'], ext=m['ext'], dur=m['dur'], tag=m['tag'], conv=m['conv'], coop=m['coop'], alt=m.get('alt'))
              for m in 주활동들 if m['block'] == b]
        if not 주활동_먼저:
            if i == 0:
                틀.append(자율(날짜))
            elif i == 1:
                틀.append(개별(날짜))
    _지정 = 하루_틀_지정.get((날짜.weekday(), '첫')) if 날짜.day <= 7 and (날짜.weekday(), '첫') in 하루_틀_지정 else 하루_틀_지정.get(날짜.weekday())
    if _지정:                              # 요일별로 정해 둔 순서대로 (마포 주간)
        틀 = []
        for 항 in _지정:
            글_, 길이_ = 항[0], 항[1]
            외부_ = bool(항[2]) if len(항) > 2 else False
            if isinstance(글_, str) and 글_.startswith('#'):
                번호_ = int(글_[1:])
                for m in [x for x in 주활동들 if x['block'] == 번호_]:
                    틀.append(dict(kind='주', text=m['text'], ext=m['ext'], dur=m['dur'], tag=m['tag'], conv=m['conv'], coop=m['coop'], alt=m.get('alt')))
            else:
                틀.append(dict(kind='보조', text=글_, ext=외부_, dur=길이_, tag=('연장' if 길이_ >= 9 else ('산책' if 외부_ else '일상')), conv=False, coop=any(k in 글_ for k in 협력_글자)))
        최소 = 10 ** 6
    elif 하루_틀_시간:
        pass
    elif 주활동_먼저:
        틀 += [자율(날짜), 개별(날짜)]
        틀.append(산책())
    else:
        if not 블록들:
            틀.append(자율(날짜))
        if len(블록들) < 2:
            틀.append(개별(날짜))
        틀.append(산책())

    누적 = 주끝 = 0
    for s in 틀:
        누적 += s['dur']
        if s['kind'] == '주':
            주끝 = 누적
    if not 하루_틀_시간 and 최소 < 주끝 - 1e-9:                      # 짧게 있는 아동도 주활동을 모두 하도록 주활동을 앞으로
        틀 = [s for s in 틀 if s['kind'] == '주'] + [s for s in 틀 if s['kind'] != '주']
    슬롯, t, 끝 = [], 시작, 시작 + 최대
    for s in 틀:
        if t >= 끝 - 1e-9:
            break
        길이 = min(s['dur'], 끝 - t)
        슬롯.append(dict(start=t, dur=길이, text=s['text'], ext=s['ext'], tag=s['tag'], conv=s['conv'], coop=s.get('coop', False), alt=s.get('alt')))
        t += 길이
    k = 0
    while t < 끝 - 1e-9:                          # 틀보다 오래 있는 날은 보조활동을 덧붙임
        글, 종류 = 덧붙일_활동[k % len(덧붙일_활동)]
        길이 = min(1, 끝 - t)
        슬롯.append(dict(start=t, dur=길이, text=글, ext=False, tag=종류, conv=True, coop=False))
        t += 길이
        k += 1
    if 같은_활동_합치기:
        합친 = []
        for x in 슬롯:
            if 합친 and 합친[-1]['text'] == x['text'] and 합친[-1]['tag'] == x['tag'] and 합친[-1].get('alt') == x.get('alt') and abs(합친[-1]['start'] + 합친[-1]['dur'] - x['start']) < 1e-9:
                합친[-1]['dur'] += x['dur']
            else:
                합친.append(dict(x))
        슬롯 = 합친
    return 슬롯


def 토요일_공통(날짜, 최대, tt, 주의, 시작=None):
    if 날짜.weekday() == 6:
        주의.append((날짜, '일요일 결제예요. 토요일 방식으로 만들었어요'))
    장소 = 글자고치기(토요일_장소들.get(f'{날짜.month}/{날짜.day}') or tt.토요일_장소(날짜))
    if not 장소:
        주의.append((날짜, f'시간표에 토요일 활동 장소가 없어 "{토요일_이름}"으로만 적었어요'))
    if 시작 is None:
        시작 = 토요일_시작시각
    if 토요일_통째외부:                             # 토요일 결제 시간 전체를 외부활동 한 칸으로
        if 장소 and 토요일_글_모양:
            글 = 토요일_글_모양.format(이름=토요일_이름, 장소=장소)
        else:
            글 = 활동글(토요일_이름, 장소) if 장소 else 토요일_이름
        return [dict(start=시작, dur=최대, text=글, ext=True, tag='토외부', conv=False, coop=False)]
    실내 = '실내' in 장소
    끝 = 시작 + 최대
    구성 = []                                      # (길이, 글, 외부, 종류)
    구성.append((토요일_외부시작 - 시작, '자율활동 트램펄린', False, '토자율'))       # 10시 전: 자율활동
    if 실내:                                      # 실내요리 날: 마트 → 요리 → 점심 → 뒷마무리
        구성 += [(1, '지역사회활동(마트 장보기)', True, '토마트'),
                (1, f'요리활동({장소})', False, '토요리'),
                (1, '점심식사', False, '토식사'),
                (1, '일상생활훈련(뒷마무리)', False, '토뒷정리')]
        뒤 = [('태블릿학습게임', '토태블릿'), ('자율활동 빈백', '토자율')]
    else:
        구성.append((토요일_외부끝 - 토요일_외부시작, f'외부활동({장소})' if 장소 else '외부활동', True, '토외부'))
        뒤 = [('태블릿학습게임', '토태블릿'), ('자율활동 빈백', '토자율')]
    슬롯, t = [], 시작
    for 길이, 글, 외부, 종류 in 구성:
        if t >= 끝 - 1e-9:
            break
        길이 = min(길이, 끝 - t)
        슬롯.append(dict(start=t, dur=길이, text=글, ext=외부, tag=종류, conv=False, coop=False))
        t += 길이
    k = 0
    while t < 끝 - 1e-9:                          # 그 밖의 시간: 보조활동(태블릿학습게임/자율활동)
        글, 종류 = 뒤[k % len(뒤)]
        길이 = min(1, 끝 - t) if 실내 else 끝 - t
        슬롯.append(dict(start=t, dur=길이, text=글, ext=False, tag=종류, conv=False, coop=False))
        t += 길이
        k += 1
    return 슬롯


def 자르기(슬롯들, 시작, 시간):
    끝 = 시작 + 시간
    결과 = []
    for s in 슬롯들:
        a, b = max(s['start'], 시작), min(s['start'] + s['dur'], 끝)
        if b - a > 1e-9:
            결과.append(dict(s, start=a, dur=b - a))
    return 결과


def 슬롯_글(슬롯들):
    if 붙은_같은_활동_합쳐쓰기:
        합 = []
        for x in 슬롯들:
            if 합 and 합[-1]['text'] == x['text'] and 합[-1].get('tag') == x.get('tag') and abs(합[-1]['start'] + 합[-1]['dur'] - x['start']) < 1e-9:
                합[-1] = dict(합[-1], dur=합[-1]['dur'] + x['dur'])
            else:
                합.append(x)
        슬롯들 = 합
    return 활동_구분.join(f'{시각(s["start"])}~{시각(s["start"] + s["dur"])} {s["text"]}' for s in 슬롯들)


def 합계(슬롯들):
    """(내부, 외부, 협력) 시간"""
    협력 = sum(s['dur'] for s in 슬롯들 if s.get('coop'))
    내부 = sum(s['dur'] for s in 슬롯들 if not s['ext'] and not s.get('coop'))
    외부 = sum(s['dur'] for s in 슬롯들 if s['ext'])
    return 내부, 외부, 협력


def 목표_시간(총, 기준):
    """총·기준 시간의 30%를 '넘는' 가장 작은 30분 단위 시간"""
    def 넘는최소(x):
        if 외부_최소비율 <= 0:          # 0이면 30% 맞추기를 하지 않아요
            return 0
        return (math.floor(x * 외부_최소비율 * 2 + 1e-9) + 1) / 2
    return 넘는최소(기준), 넘는최소(총)


def 외부_맞추기(날정보, 아동시간, 기준, 아동시작=None, 제외=None):
    """모든 아동의 외부+협력 시간이 30%를 넘도록, 공통 시간표의 보조활동(자율/일상/태블릿/개별)을 산책으로 바꿉니다.
    반환: {아동: (총, 외부, 목표, 기준가능)}"""
    목표, 총합 = {}, {}
    for n, 일들 in 아동시간.items():
        총 = sum(일들.values())
        g기준, g총 = 목표_시간(총, 기준[n] if isinstance(기준, dict) else 기준)
        목표[n] = (g기준 if g기준 <= 총 else g총, g기준 <= 총)
        총합[n] = 총

    def 시작(n, d):
        return (아동시작 or {}).get(n, {}).get(d, 날정보[d]['start'])

    def 몇개(n, d):
        글 = (제외 or {}).get((n, d))
        선택 = (제외 or {}).get(('선택', n, d))
        슬롯들 = [dict(x, **x['alt']) if (선택 and x.get('alt')) else x for x in 날정보[d]['slots']]
        return [x for x in 슬롯들 if not (글 and 글 in x['text'] and not x.get('alt'))]

    def 외부시간(n):
        return sum(sum(합계(자르기(몇개(n, d), 시작(n, d), h))[1:]) for d, h in 아동시간[n].items())

    while True:
        부족 = [n for n in 아동시간 if 외부시간(n) < 목표[n][0] - 1e-9]
        if not 부족:
            break
        best = None
        for d, 정보 in sorted(날정보.items()):
            if d.weekday() >= 5:
                continue
            for s in 정보['slots']:
                if not s['conv'] or s['ext'] or s.get('coop'):
                    continue
                이득 = sum(sum(x['dur'] for x in 자르기([s], 시작(n, d), 아동시간[n][d])) for n in 부족 if d in 아동시간[n])
                if 이득 <= 1e-9:
                    continue
                키 = (-이득, 바꿀_차례.index(s['tag']), 정보['바꿈수'], d, s['start'])
                if best is None or 키 < best[0]:
                    best = (키, 정보, s)
        if not best:
            break
        _, 정보, s = best
        s.update(text='산책', ext=True, tag='산책', conv=False, coop=False)
        정보['바꿈수'] += 1
    return {n: (총합[n], 외부시간(n), 목표[n][0], 목표[n][1]) for n in 아동시간}


def 산책_장소_정하기(날정보):
    """산책하는 날마다 중랑천 → 묵동어린이공원 → 화랑대 철도공원 순으로 번갈아 갑니다."""
    순번 = 0
    for d in sorted(날정보):
        산책들 = [s for s in 날정보[d]['slots'] if s['tag'] == '산책']
        if 산책들 and 산책_장소들:
            장소 = 산책_장소들[순번 % len(산책_장소들)]
            순번 += 1
            for s in 산책들:
                s['text'] = f'산책 {장소}'


# =====================================================================
# 4. 현양식에 채우기
# =====================================================================
첫날줄 = 10        # 양식에서 날짜 줄이 시작하는 줄
견본줄 = 11        # 날짜 줄의 모양을 가져올 줄
끝줄원본 = 37      # 양식의 '합계' 줄 (아래 안내문까지 44줄)
마지막원본 = 44


def 레이아웃_찾기(ws):
    """양식의 글자(활동계획서 / 날짜 / 합계)로 줄 위치를 알아냅니다. 양식 위쪽에 줄이 더 있어도 따라갑니다."""
    제목 = 머리 = 합계 = None
    for r in range(1, min(ws.max_row, 80) + 1):
        v = ws.cell(r, 1).value
        if not isinstance(v, str):
            continue
        t = v.replace(' ', '').replace('\n', '')
        if 제목 is None and '활동계획서' in t and '활동기록지' in t:
            제목 = r
        elif 머리 is None and t == '날짜':
            머리 = r
        elif 합계 is None and t == '합계' and 머리 is not None:
            합계 = r
    위 = (제목 - 1) if 제목 else 0
    첫날 = (머리 + 1) if 머리 else 첫날줄 + 위
    끝 = 합계 or (끝줄원본 + 위)
    견본 = 첫날 + 1 if 첫날 + 1 < 끝 else 첫날
    return dict(위=위, 첫날=첫날, 견본=견본, 끝=끝, 마지막=max(ws.max_row, 끝))


def 칸_옆값(ws, 글자, 위=0):
    """왼쪽 제목 칸(예: '수급시간')의 오른쪽 칸 값을 돌려줍니다. 못 찾으면 None."""
    for r in range(1, min(ws.max_row, 40) + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.replace(' ', '').replace('\n', '') == 글자:
            return ws.cell(r, 4).value
    return None


def 스냅샷(ws, 마지막원본=마지막원본):
    snap = dict(style={}, value={}, height={})
    for r in range(1, 마지막원본 + 1):
        for c in range(1, 16):
            cell = ws.cell(r, c)
            snap['style'][(r, c)] = copy(cell._style)
            snap['value'][(r, c)] = cell.value
        snap['height'][r] = ws.row_dimensions[r].height
    return snap


def 줄_모양(ws, r, snap, 견본):
    for c in range(1, 16):
        ws.cell(r, c)._style = copy(snap['style'][(견본, c)])


def 시트_채우기(ws, 정보, 계획서):
    레 = 레이아웃_찾기(ws)
    위, 첫날, 견본, 끝줄, 마지막 = 레['위'], 레['첫날'], 레['견본'], 레['끝'], 레['마지막']
    snap = 스냅샷(ws, 마지막)
    아동, 월 = 정보['아동'], 정보['월']
    작성일 = 정보['계획일'] if 계획서 else 정보['기록일']

    for m in list(ws.merged_cells.ranges):               # 날짜 줄 아래쪽 병합 모두 풀기
        if m.min_row >= 첫날:
            ws.unmerge_cells(str(m))
    for r in range(첫날, 첫날 + 70):
        for c in range(1, 16):
            ws.cell(r, c).value = None
            ws.cell(r, c)._style = StyleArray()
        ws.row_dimensions[r].height = None

    def 요구사항_줄(행, 글):
        줄_모양(ws, 행, snap, 견본)
        ws.merge_cells(start_row=행, start_column=1, end_row=행, end_column=15)
        칸 = ws.cell(행, 1)
        칸.value = 글
        칸.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True, indent=1)
        칸.font = Font(name='맑은 고딕', size=11, color='FF000000')
        ws.row_dimensions[행].height = max(39.95, 줄수(글, 150) * 16.5 + 12)

    행 = 첫날
    for 일 in (정보['날들'][::-1] if (not 계획서 and 정보['아동'] in 기록지_역순) else 정보['날들']):
        줄_모양(ws, 행, snap, 견본)
        for (a, b) in ((4, 6), (7, 9), (10, 11)):
            ws.merge_cells(start_row=행, start_column=a, end_row=행, end_column=b)
        ws.cell(행, 1).value = 일['날짜'].day
        ws.cell(행, 2).value = 요일글자[일['날짜'].weekday()]
        ws.cell(행, 3).value = 일['인력']
        if 일['활동']:
            ws.cell(행, 4).value = 일['활동']
            if not 계획서:                               # 계획서 시트에는 활동기록을 넣지 않음
                ws.cell(행, 7).value = 일['활동']
        특이 = 일['특이_계획'] if 계획서 else 일['특이_기록']
        ws.cell(행, 10).value = 특이 or None
        if 일['활동']:
            값들 = (일['내부'], 일['외부'], 일.get('협력', 0), 일.get('송영', 0))
            for c, v in zip((12, 13, 14, 15), 값들):
                if v or 오른쪽_0_쓰기:
                    ws.cell(행, c).value = 숫자(v)
        for c in range(1, 16):
            검정(ws.cell(행, c))
        ws.row_dimensions[행].height = max(
            39.95, 줄수(일['활동'], 47) * 16.5 + 8, 줄수(특이, 21) * 16.5 + 8)
        행 += 1
        if not 요구사항_맨아래 and not 요구사항_특이사항 and 정보["요구_줄날짜"] == 일["날짜"] and not (빈_요구사항_줄_빼기 and not 정보['요구사항']):   # 인정결제 줄 바로 아래: 병합된 요구사항 칸
            요구사항_줄(행, 정보['요구사항'])
            행 += 1
    if 요구사항_맨아래:                                      # 날짜 줄들 맨 아래에 한 줄
        조각 = [x for x in (정보['요구사항'], (None if 계획서 else 정보.get('인정글'))) if x]
        요구사항_줄(행, 요구사항_머리말 + ' / '.join(조각))
        행 += 1
    마지막날 = 행 - 1

    # 합계 줄 + 아래 안내문/확인란 (양식의 합계 줄부터 끝까지를 그대로 옮김)
    합계행 = 행
    마지막행 = 합계행
    for 원 in range(끝줄, 마지막 + 1):
        내용 = snap['value'][(원, 1)]
        if isinstance(내용, str) and any(k in 내용 for k in 안내문_지울_글자):
            continue
        r = 합계행 + (원 - 끝줄)
        for c in range(1, 16):
            ws.cell(r, c)._style = copy(snap['style'][(원, c)])
        ws.cell(r, 1).value = 내용
        ws.row_dimensions[r].height = snap['height'][원]
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11 if 원 == 끝줄 else 15)
        마지막행 = r
    for c in range(12, 16):
        Lc = get_column_letter(c)
        ws.cell(합계행, c).value = f'=SUM({Lc}{첫날}:{Lc}{마지막날})'
        검정(ws.cell(합계행, c))
    _제공 = _이용 = None                                    # 서명 줄 위치는 글자로 찾아요 (마포점 양식은 '제공기관으로부터…' 안내 줄이 하나 더 있어요)
    for _r in range(합계행, 마지막행 + 1):
        _t = str(ws.cell(_r, 1).value or '').replace(' ', '')
        if _제공 is None and _t.startswith('제공기관명'):
            _제공 = _r
        elif _이용 is None and _t.startswith('이용자(혹은'):
            _이용 = _r
    _제공 = _제공 or (합계행 + 5)
    _이용 = _이용 or (합계행 + 6)
    _날짜줄 = _제공 - 1
    _원날짜 = snap['value'].get((끝줄 + (_날짜줄 - 합계행), 1))
    if 날짜칸_날짜형 or isinstance(_원날짜, (dt.datetime, dt.date)):
        ws.cell(_날짜줄, 1).value = dt.datetime(작성일.year, 작성일.month, 작성일.day)
        if ws.cell(_날짜줄, 1).number_format == 'General':
            ws.cell(_날짜줄, 1).number_format = 'yyyy"년" m"월" d"일"'
        ws.cell(_날짜줄, 1).alignment = Alignment(horizontal='center', vertical='center')
    else:
        ws.cell(_날짜줄, 1).value = f'{작성일.year} 년   {작성일.month} 월   {작성일.day} 일'          # 41줄
    ws.cell(_제공, 1).value = f'제공기관 명   {기관명}   (인)'                                 # 42줄
    ws.cell(_이용, 1).value = f'이용자(혹은  보호자)명   {아동}   (인)'                        # 43줄
    for i in (_날짜줄, _제공, _이용):
        검정(ws.cell(i, 1))

    # 윗부분 (양식 위쪽에 줄이 더 있으면 그만큼 아래로 따라갑니다)
    def 칸(r, c):
        return ws.cell(r + 위, c)
    제목 = re.sub(r'\(\s*\)', f'( {월} )', str(snap['value'][(1 + 위, 1)]))
    if 계획서:
        제목 = 제목.replace('□활동계획서', '■활동계획서').replace('■활동기록지', '□활동기록지')
    else:
        제목 = 제목.replace('□활동기록지', '■활동기록지').replace('■활동계획서', '□활동계획서')
    칸(1, 1).value = 제목
    칸(2, 4).value = 정보['담당']                  # 작성자 = 그 아동의 담당 선생님
    칸(2, 10).value = 날짜글(작성일)
    칸(3, 4).value = 아동
    칸(4, 5).value = 정보['담당']
    칸(4, 7).value = ', '.join(정보['대체인력']) if 정보['대체인력'] else None
    for c in (5, 7):
        칸(4, c).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    if 수급시간_칸_채우기:
        칸(5, 4).value = f'■ 월 {시간글자(정보["기준"])}시간' + (' (개인예산제 사업으로 수급시간 단축)' if 정보.get('개인예산') else '')
    칸(6, 4).value = f'월 (   {시간글자(정보["총"])}   )시간'
    _사유 = 축소사유_지정.get(정보['아동'], 계획시간_축소사유)
    칸(7, 4).value = _사유 if 정보['총'] < 정보['기준'] else None
    if 정보['총'] < 정보['기준'] and '\n' in str(_사유 or ''):          # 여러 줄이면 줄 높이를 늘림
        _r = 칸(7, 4).row
        ws.row_dimensions[_r].height = max(ws.row_dimensions[_r].height or 15, 15 * (str(_사유).count('\n') + 1) + 3)
    for k in range(4):
        칸(4 + k, 10).value = f'=({get_column_letter(12 + k)}{합계행})'
    if 정보['기준'] != 66:                         # 양식에 66 으로 박혀 있는 비율 계산(66시간, =J4/66)을 그 아동 수급시간으로
        기 = 시간글자(정보['기준'])
        for r in range(3, 8):
            v = 칸(r, 12).value
            if isinstance(v, str) and '66' in v:
                칸(r, 12).value = v.replace('66', 기)
    for r, c in ((2, 4), (3, 4), (4, 5), (4, 7), (5, 4), (6, 4), (7, 4), (2, 10)):
        검정(칸(r, c))
    ws.print_area = f'A1:O{마지막행}'


# =====================================================================
# 지점 설정 파일 (지점설정.txt) - 코드를 고치지 않고 메모장으로 지점별 값을 바꿉니다
# =====================================================================
설정폴더들 = []
설정_이름들 = ['내_이름', '기관명', '기본_시작시각', '늦은출근_선생님', '토요일_시작시각', '토요일_외부시작', '토요일_외부끝',
            '산책_장소들', '난타_시간', '난타_인원', '외부_최소비율', '시트이름_구분', '대체인력_문구', '계획시간_축소사유', '축소사유_지정',
            '모든_아동', '아동_직접지정', '추가_공휴일', '외부활동_글자', '활동_길이', '활동_뒤', '글자_바꾸기', '이름_바꾸기',
            # --- 노원점 등에서 쓰는 옵션 ---
            '결제시각으로_시작', '시각_반올림_분', '평일_가장이른시각', '시각_표기', '활동_구분', '협력_글자', '시간대_줄표시', '상세로_대체', '속_줄바꿈_펴기', '연장_끝까지', '기록지_역순', '빈_요구사항_줄_빼기', '상세만_쓰는_글자',
            '같은_활동_합치기', '붙은_같은_활동_합쳐쓰기', '같은_주활동_하루_한번', '주활동_먼저', '산책_가능한_주활동', '자율활동_글', '자율활동_지정', '개별학습_글', '덧붙일_활동',
            '사업_글자', '대상_서비스', '하루_틀_지정', '시간표_시트', '토요일_통째외부', '토요일_이름', '토요일_장소들', '모든_운영일_행', '인정결제만_있는_날_줄_빼기', '오른쪽_0_쓰기', '안내문_지울_글자', '요구사항_맨아래',
            '요구사항_머리말', '인정결제_요구사항에', '기록일_방식', '소급_문구', '대체_기록지에만', '수급시간_바꾸기', '일부만_하는_활동', '일부만_대신_활동',
            '수급시간_칸_채우기', '빈내용_알림', '양천식_글', '소제목_빼기', '시각_올림', '담당_지정',
            '송영_글에_쓰기', '송영_따로_처리', '공통_시작_고정', '하루_틀_시간', '보조_맨앞_요일', '협력_먼저', '연장_활동', '활동_연결',
            '토요일_글_모양', '날짜_모양', '날짜칸_날짜형', '인정결제_문구', '인정결제_덧붙임', '인정결제_표기', '요구사항_특이사항', '요구사항_자동']
설정_설명 = {
    '내_이름': '결제내역에 적힌 내 이름 (모든_아동 = True 이면 크게 중요하지 않아요)',
    '기관명': '기관(지점) 이름',
    '기본_시작시각': '평일 활동 시작 시각 (14 = 오후 2시)',
    '늦은출근_선생님': "늦게 시작하는 선생님과 시작 시각. 예: {'홍길동': 15, '김철수': 15}   (없으면 {} )",
    '토요일_시작시각': '토요일 활동 시작 시각',
    '토요일_외부시작': '토요일 외부활동 시작 시각',
    '토요일_외부끝': '토요일 외부활동 끝 시각',
    '산책_장소들': "산책 장소(날마다 번갈아 감). 예: ['한강공원', '근처 놀이터']",
    '난타_시간': '난타 시간(시간). 난타가 없는 지점은 그대로 두어도 돼요',
    '난타_인원': '난타 하는 총 인원',
    '시트이름_구분': "시트 이름 가운데 글자. 기본 ' ' (구교현 계획서), 광진점 예시는 '_' (구교현_계획서)",
    '외부_최소비율': '외부+협력 활동이 수급시간의 이 비율을 넘어야 해요 (0.30 = 30%)',
    '대체인력_문구': "대체 선생님이 들어간 날 개별 특이사항에 적을 말. 대체 선생님 이름이 들어갈 자리는 {대체} 로 쓰세요",
    '계획시간_축소사유': '총 계획시간이 수급시간보다 적을 때 적을 말',
    '축소사유_지정': "아동마다 다른 계획시간 축소 사유. 예: {'김하나': '- 치료 일정\\n- 학원 일정'}  (없으면 {} )",
    '모든_아동': 'True = 결제내역의 모든 아동 / False = 내_이름 선생님 담당 아동만',
    '아동_직접지정': "특정 아동만 만들고 싶을 때. 예: ['김하나', '이둘']  (비우면 [] )",
    '추가_공휴일': "임시공휴일 등. 예: ['2027-06-03']  (보통 [] )",
    '외부활동_글자': "활동 이름에 이 글자가 들어 있으면 센터 밖 활동으로 봐요. 예: ['지역사회', '견학']",
    '활동_길이': "활동 이름에 이 글자가 들어 있으면 그 시간만큼 해요(기본 1시간). 예: {'난타': 2, '수영': 2}",
    '활동_뒤': "활동 글 뒤에 붙일 말. 예: {'난타': '(총 8명)'}  (없으면 {} )",
    '글자_바꾸기': "시간표 글자를 기록지용 글자로 바꿔요. 예: {'마트장보기': '마트 장보기'}",
    '이름_바꾸기': "활동 이름을 바꿔요. 예: {'지역사회': '지역사회활동'}",
    '결제시각으로_시작': 'True = 결제내역의 시작 시각(반올림)부터 활동을 시작해요. False = 위의 기본_시작시각 사용',
    '시각_반올림_분': '시작 시각을 몇 분 단위로 맞출지. 30이면 14:45~15:14 → 15:00, 15:15~15:44 → 15:30',
    '평일_가장이른시각': '평일 활동이 이 시각보다 일찍 시작하지 않아요. 예: 15 (0이면 제한 없음)',
    '시각_표기': "'콜론' → 15:30~16:00 / '시' → 15.5시~16시 / '시분' → 15시 30분~16시",
    '활동_구분': "활동 사이를 줄바꿈('\\n')으로 할지, ' / ' 로 가로 한 줄로 쓸지",
    '협력_글자': "활동 글에 이 글자가 있으면 협력기관 시간으로 세요. 예: ['(협)', '볼링장']",
    '빈_요구사항_줄_빼기': 'True = 요구사항이 비어 있으면 요구사항 빈 줄을 만들지 않아요',
    '기록지_역순': "기록지만 날짜를 말일→1일 순서로 쓸 아동. 예: ['유효린', '허서현']",
    '시간대_줄표시': "True = 시간표 첫 칸에 '15 - 16시' 처럼 시간대가 적힌 줄을 활동 줄로 읽어요",
    '상세로_대체': 'True = 시간표 세부내용이 있으면 활동 이름 대신 세부내용만 써요',
    '속_줄바꿈_펴기': 'True = 시간표 한 칸 안의 줄바꿈을 띄어쓰기로 바꿔요',
    '연장_끝까지': 'True = 연장 활동을 결제 끝 시각까지 이어서 써요 (예: 18시 이후 전부 일상생활)',
    '상세만_쓰는_글자': "시간표 세부내용에 이 글자가 있으면 활동 이름 없이 세부내용만 써요. 예: ['(협)']",
    '붙은_같은_활동_합쳐쓰기': 'True = 붙어 있는 같은 활동(산책 등)은 17:00~19:00 산책 처럼 한 줄로 적어요. False = 1시간씩 따로',
    '같은_주활동_하루_한번': 'True = 같은 주활동이 하루에 두 번 시간표에 있어도 한 번만 넣어요',
    '같은_활동_합치기': 'True = 붙어 있는 같은 활동은 한 칸으로 합쳐요 (외부활동 A+B 등)',
    '주활동_먼저': 'True = 시간표 활동을 모두 먼저 하고, 자율활동 등은 남는 시간에 해요',
    '산책_가능한_주활동': "30%가 모자랄 때 산책으로 바꿔도 되는 시간표 활동. 예: ['일상생활']",
    '자율활동_글': '자율활동 칸에 적을 말. 예: 자율활동(캐치볼)',
    '자율활동_지정': "아이마다 다른 자율활동. 예: {'김하나': ['자율활동(캐치볼)', '자율활동(영단어)'], '이둘': '자율활동(한강라면)'}  (없으면 {} )",
    '개별학습_글': '개별학습 칸에 적을 말',
    '덧붙일_활동': "시간이 남을 때 차례로 덧붙일 활동. 예: [('자율활동', '자율'), ('태블릿학습게임', '태블릿'), ('개별학습', '개별')]  (뒤의 낱말은 자율/일상/태블릿/개별 중 하나)",
    '사업_글자': "결제내역에 사업이 섞여 있을 때 이 글자가 든 사업만 써요. 예: '발달장애인 주간'",
    '대상_서비스': "결제내역에서 만들 서비스. 방과후='방과후활동서비스', 주간='주간활동서비스'",
    '하루_틀_지정': "요일별 하루 순서를 직접 정해요 (마포 주간). '#N'은 시간표 N번째 묶음",
    '시간표_시트': "시간표 엑셀에 시트가 여러 개일 때 쓸 시트 이름. 예: '방과후시간표'",
    '토요일_통째외부': 'True = 토요일은 결제 시간 전체를 외부활동 한 칸으로 적어요',
    '토요일_이름': "토요일 활동 이름. 예: '토요외부활동'",
    '토요일_장소들': "날짜별 토요일 장소(시간표 대신). 예: {'9/5': '고양 낙농치즈체험관'}",
    '모든_운영일_행': 'True = 결제가 없는 날도 (일요일·공휴일 빼고) 날짜 줄을 만들어요',
    '인정결제만_있는_날_줄_빼기': 'True = 인정결제만 있고 활동이 없는 날은 날짜 줄을 만들지 않아요',
    '오른쪽_0_쓰기': 'True = 내부/외부/협력/송영 칸에 0도 써요',
    '안내문_지울_글자': "양식 아래 안내문 중 이 글자가 있는 줄은 지워요. 예: ['210㎜']",
    '요구사항_자동': 'True = 시간표의 요구사항 문구를 프로그램이 써요 (중랑). False = 쓰지 않고 비워 둬요 (다른 지점은 직접 적어요)',
    '요구사항_맨아래': 'True = 요구사항을 날짜 줄들 맨 아래에 한 줄(이용자 요구사항) ...)로 써요',
    '요구사항_머리말': "요구사항 줄 맨 앞 글자. 예: '이용자 요구사항) '",
    '인정결제_요구사항에': 'True = 인정결제 시간을 요구사항 줄 끝에 쓰고, 특이사항에는 쓰지 않아요 (기록지만)',
    '기록일_방식': "기록지 작성일자: '인정결제' = 인정결제된 날 / '말일' = 그 달 마지막 근무일",
    '소급_문구': "결제 사유 → 특이사항 문구. 예: {'대상자 과실(카드 미소지 등)': '카드 미소지로 소급결제'}",
    '대체_기록지에만': 'True = 대체인력 안내를 기록지에만 적어요',
    '일부만_하는_활동': "시간표 활동 중 몇 명만 하는 활동. 예: {'카엘': 4} (글자가 들어간 활동을 그 날 랜덤 4명만, 나머지는 자율활동). 없으면 {}",
    '일부만_대신_활동': "나머지 아이가 대신 하는 활동 글. 비우면 자율활동",
    '수급시간_바꾸기': "개인예산제 등으로 수급시간이 다른 아동. 예: {'김하나': 52.5}  (없으면 {} )",
    '송영_글에_쓰기': 'False = 송영서비스를 활동 글에는 쓰지 않고 송영 칸에만 세요',
    '송영_따로_처리': "True = 결제내역의 송영서비스를 활동과 따로 세고, 송영 칸(O)과 '송영서비스' 시간으로 써요",
    '공통_시작_고정': 'True = 같은 반은 날마다 기본 시작 시각(예: 15시)부터 같은 표를 쓰고, 아동은 자기 결제 시각 이후만 써요',
    '하루_틀_시간': '하루 기본 틀 시간 (예: 4 → 15~19시). 그 뒤는 연장 활동. 0이면 사용 안 함',
    '보조_맨앞_요일': '자율활동을 맨 앞에 두는 요일. 0=월 1=화 2=수 3=목 4=금 5=토. 예: [0, 1]',
    '협력_먼저': 'True = 협력기관 활동을 주활동 중 가장 먼저 해요',
    '연장_활동': "하루 기본 틀 뒤 첫 시간 활동. 예: {0: '요리활동-스파게티', 1: '요리활동-야채볶음밥'}  (없으면 {} )",
    '활동_연결': "활동 이름과 세부내용 연결: '(' → 미술활동(가전제품색칠) / '-' → 미술활동-가전제품색칠",
    '토요일_글_모양': "토요일 활동 글 모양. 예: '{이름} ({장소})' (비우면 기본)",
    '날짜_모양': "'붙임' → 2026.8.31. (월) 모양으로 써요 (비우면 기본)",
    '날짜칸_날짜형': 'True = 맨 아래 작성일을 글자가 아니라 날짜 칸으로 써요',
    '인정결제_문구': "인정결제 줄에 적을 말. {n}에 시간이 들어가요. 예: '인정결제 : {n} 시간'",
    '인정결제_덧붙임': "인정결제 줄 아래에 한 줄 더 적을 말. 예: '이용자 요구사항 : 원함'",
    '인정결제_표기': 'False = 인정결제 시간을 어디에도 적지 않아요',
    '요구사항_특이사항': "True = 요구사항을 기록지 마지막 날짜 줄의 개별 특이사항에 '요구사항: ...' 로 써요",
    '시각_올림': "True = 결제 시작 시각을 올림해서 활동 시작 시각으로 써요 (16:14 → 16:30)",
    '양천식_글': "True = 대제목(소제목) 글을 양천 손기록지 모양으로 적어요",
    '소제목_빼기': "이 글자가 든 활동은 세부내용 없이 이름만 적어요. 예: ['스마트교육센터']",
    '빈내용_알림': "False = 시간표에 세부내용이 비어 있어도 점검표에 알리지 않아요",
    '담당_지정': "담당 선생님을 직접 정해요. 예: {'김서진': '손은주'}  (없으면 {} → 결제를 가장 많이 한 선생님)",
    '수급시간_칸_채우기': "True = 양식의 '수급시간' 칸을 ■ 월 66시간 으로 채워요",
}


def 설정_견본_만들기(폴더):
    """지점설정.txt 가 없으면 현재 값으로 견본을 만들어 줍니다. 경로를 돌려줍니다."""
    경로 = os.path.join(폴더, '지점설정.txt')
    if not os.path.exists(경로):
        줄 = ['# 지점 설정 파일 - 메모장으로 열어 "=" 오른쪽 값만 고치고 저장하세요.',
              '# 줄 맨 앞에 # 이 있으면 설명(무시)이에요. 고치지 않은 항목은 프로그램 기본값을 써요.',
              '# 글자는 \'작은따옴표\'로 감싸고, 여러 개는 [ ] 안에 쉼표로 나누세요.', '']
        g = globals()
        for 이름 in 설정_이름들:
            줄.append('# ' + 설정_설명.get(이름, ''))
            줄.append(f'{이름} = {g[이름]!r}')
            줄.append('')
        with open(경로, 'w', encoding='utf-8-sig') as f:
            f.write('\n'.join(줄))
    return 경로


def 설정_불러오기():
    """지점설정.txt 가 있으면 읽어서 설정값을 바꿉니다. 바뀐 항목 이름 목록과 문제 목록을 돌려줍니다."""
    import ast
    폴더들 = list(dict.fromkeys([p for p in (설정폴더들 or []) + 기준폴더들() if p]))
    g = globals()
    바뀜, 문제 = [], []
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
                문제.append(f'지점설정.txt {번호}번째 줄: "{이름}" 은(는) 없는 항목이에요')
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
    if '난타_시간' in 바뀜 or '난타_인원' in 바뀜:
        if '활동_길이' not in 바뀜:
            g['활동_길이'] = dict(g['활동_길이'], **{'난타': g['난타_시간']})
        if '활동_뒤' not in 바뀜:
            g['활동_뒤'] = dict(g['활동_뒤'], **{'난타': f"(총 {g['난타_인원']}명)"})
    return 바뀜, 문제


# =====================================================================
# 5. 전체 실행
# =====================================================================
def main():
    global 연도
    바뀜, 설정문제 = 설정_불러오기()
    if 바뀜:
        print('지점설정.txt 를 적용했어요:', ', '.join(바뀜))
    for 글 in 설정문제:
        print('※', 글)
    결제경로 = 파일_고르기(결제내역_파일, '결제')
    if not 결제경로:
        raise SystemExit('결제내역 엑셀을 찾지 못했어요. 이 프로그램과 같은 폴더에 결제내역 엑셀을 넣어 주세요.')
    자료 = 결제내역_읽기(결제경로)
    if not 자료:
        raise SystemExit('결제내역에서 읽을 내용이 없어요.')
    월 = Counter(r['월'] for r in 자료).most_common(1)[0][0]
    if not 연도:
        오늘 = dt.date.today()
        연도 = 오늘.year - 1 if 월 > 오늘.month else 오늘.year
    다른달 = Counter(r['월'] for r in 자료 if r['월'] != 월)
    자료 = [r for r in 자료 if r['월'] == 월]
    if 사업_글자:
        자료 = [r for r in 자료 if 사업_글자 in str(r.get('사업') or '')]
    for r in 자료:
        r['날짜'] = dt.date(연도, r['월'], r['일'])

    양식경로 = 파일_고르기(양식_파일, '양식')
    시간표경로 = 파일_고르기(시간표_파일, '시간표', 월)
    if not 양식경로:
        raise SystemExit('현양식 엑셀을 찾지 못했어요. 첫 줄에 "활동계획서 / 활동기록지"가 있는 양식 파일을 폴더에 넣어 주세요.')
    print('결제내역:', os.path.basename(결제경로))
    print('시간표  :', os.path.basename(시간표경로) if 시간표경로 else '(없음)')
    print('양식    :', os.path.basename(양식경로))

    tt = 시간표(시간표경로, 연도, 월)
    요구목록 = 요구사항_읽기(시간표경로)
    if not 시간표경로:
        print('※ 시간표 파일을 못 찾았어요. 활동 이름이 기본 글자로만 채워집니다.')
    elif not tt.ok:
        print('※ 시간표 모양을 읽지 못했어요. 활동 이름이 기본 글자로만 채워집니다.')
    elif tt.mode == 'order':
        print('※ 시간표의 주차 날짜가 이번 달과 달라서, 시간표의 첫째 주→첫째 주, 둘째 주→둘째 주 순서로 맞췄어요.')
    if not 요구목록:
        print('※ 시간표 파일에 "요구사항" 시트가 없어 요구사항 칸에는 제목만 적습니다.')

    wb0 = 엑셀열기(양식경로)
    기준 = 66
    _ws0 = wb0.worksheets[0]
    m = re.search(r'(\d+(?:\.\d+)?)\s*시간', str(_ws0.cell(5 + 레이아웃_찾기(_ws0)['위'], 4).value or ''))
    if m:
        기준 = float(m.group(1)) if '.' in m.group(1) else int(m.group(1))
    기준맵 = {}

    아동별 = defaultdict(list)
    for r in 자료:
        아동별[r['아동']].append(r)

    if 아동_직접지정:
        대상 = [a for a in 아동_직접지정 if a in 아동별]
        for a in [a for a in 아동_직접지정 if a not in 아동별]:
            print('※ 결제내역에 없는 이름:', a)
    else:
        대상 = []
        for a, rs in 아동별.items():
            주인 = Counter(x['인력'] for x in rs if x['유형'] == 대상_서비스)
            if not 주인:
                continue
            if 모든_아동 or 주인.most_common(1)[0][0] == 내_이름:
                대상.append(a)
        대상.sort()
    if not 대상:
        raise SystemExit(f'"{내_이름}" 선생님이 담당인 아동을 찾지 못했어요. 설정의 내_이름을 확인해 주세요.')

    바탕 = None
    if 저장_위치 == '바탕화면':
        for 후보 in (os.path.join(os.path.expanduser('~'), 'Desktop'),
                    os.path.join(os.path.expanduser('~'), '바탕 화면'),
                    os.path.join(os.environ.get('OneDrive', '~'), 'Desktop'),
                    os.path.join(os.environ.get('OneDrive', '~'), '바탕 화면')):
            if os.path.isdir(os.path.expanduser(후보)):
                바탕 = os.path.expanduser(후보)
                break
    폴더이름 = f'{출력_폴더}_{월}월' + (f'_{지점이름}' if 지점이름 else '')
    출력 = os.path.join(바탕 or 기준폴더들()[0], 폴더이름)
    os.makedirs(출력, exist_ok=True)
    점검 = []
    for 이름, 글 in 결제_건너뜀:
        점검.append((이름, None, 글))
    for 달, 개수 in sorted(다른달.items()):
        점검.append(('(결제내역 전체)', None, f'{달}월 결제 {개수}줄은 이번 달({월}월)이 아니라서 뺐어요'))

    # ---- 1단계: 아동별·날짜별 결제 정리 (같은 날 서비스+인정결제는 한 줄)
    아동일 = {}
    담당표 = {}
    for 아동 in 대상:
        rs = 아동별[아동]
        주인 = Counter(x['인력'] for x in rs if x['유형'] == 대상_서비스)
        담당표[아동] = 담당_지정.get(아동) or (주인.most_common(1)[0][0] if 주인 else 내_이름)
        if len(주인) >= 2 and 아동 not in 담당_지정:
            (a1, n1), (a2, n2) = 주인.most_common(2)
            if n2 >= 3:
                점검.append((아동, None, f'결제한 선생님이 {a1} {n1}일, {a2} {n2}일로 비슷해요. 중간에 담당이 바뀐 거라면 지점설정의 담당_지정에 {{"{아동}": "선생님이름"}} 을 적어 주세요 (지금은 {a1} 선생님을 담당으로 했어요)'))
        날별 = {}
        for x in sorted(rs, key=lambda z: (z['날짜'], z['유형'])):
            d = 날별.setdefault(x['날짜'], dict(날짜=x['날짜'], 시간=0.0, 인정=0.0, 인력=None, 사유=[], 시작=None, 송영=0.0, 송영시각=None))
            if 대상_서비스 != '방과후활동서비스' and x['유형'] not in ('인정결제', '송영서비스', 대상_서비스):
                continue
            if 송영_따로_처리 and x['유형'] == '송영서비스':
                d['송영'] += x['시간']
                if x.get('시각') is not None:
                    d['송영시각'] = x['시각']
                d['인력'] = d['인력'] or x['인력']
                continue
            if x['유형'] == '인정결제':
                d['인정'] += x['시간']
                d['인력'] = d['인력'] or x['인력']
                if x['사유'] or x['반납']:
                    점검.append((아동, x['날짜'], f'인정결제 줄의 결제내역 메모 확인: {x["사유"]} {x["반납"]}'.strip()))
                continue
            if d['시간'] > 0:
                점검.append((아동, x['날짜'], '같은 날 결제가 두 줄이라 시간을 합쳤어요. 확인해 주세요'))
            d['시간'] += x['시간']
            d['인력'] = x['인력']
            if x.get('시각') is not None:
                d['시작'] = x['시각'] if d['시작'] is None else min(d['시작'], x['시각'])
            if x['사유'] and x['사유'] not in d['사유']:
                d['사유'].append(x['사유'])
                if x['사유'] != '대상자 과실(카드 미소지 등)' and x['사유'] not in 소급_문구:
                    점검.append((아동, x['날짜'], f'일괄결제 사유가 특이해요: "{x["사유"]}" (소급결제 사유로 그대로 적었어요)'))
            if x['반납']:
                점검.append((아동, x['날짜'], f'반납구분 "{x["반납"]}" 이(가) 있어요. 확인해 주세요'))
        if 모든_운영일_행:                               # 결제가 없는 날도 (일요일·공휴일 빼고) 날짜 줄을 만듭니다
            for k in range(1, calendar.monthrange(연도, 월)[1] + 1):
                날 = dt.date(연도, 월, k)
                if not 쉬는날(날) and 날 not in 날별:
                    날별[날] = dict(날짜=날, 시간=0.0, 인정=0.0, 인력=None, 사유=[], 시작=None, 송영=0.0, 송영시각=None)
        아동일[아동] = 날별
        기준맵[아동] = 수급시간_바꾸기.get(아동, 기준)

    # ---- 2단계: 선생님별로, 날짜마다 '공통 시간표' 1개 (같은 선생님 아동은 같은 시간에 같은 활동)
    아동시간 = {a: {d: v['시간'] for d, v in 아동일[a].items() if v['시간'] > 0} for a in 대상}
    내날정보, 결과비교, 그룹들 = {}, {}, []
    for 선생님 in sorted({담당표[a] for a in 대상}):
        모둠 = [a for a in 대상 if 담당표[a] == 선생님]
        시작시각 = 늦은출근_선생님.get(선생님, 기본_시작시각)
        날정보 = {}
        for d in sorted({d for a in 모둠 for d in 아동시간[a]}):
            시간들 = [아동시간[a][d] for a in 모둠 if d in 아동시간[a]]
            주말 = d.weekday() >= 5
            시작들 = []                       # 결제 시각 기준: 아동마다 (반올림한) 시작 시각, 아니면 정해진 시작 시각
            for a in 모둠:
                if d not in 아동시간[a]:
                    continue
                원시작 = 아동일[a][d].get('시작')
                if 결제시각으로_시작 and 원시작 is not None:
                    시각분 = 시각_반올림_분
                    v = (math.ceil((원시작 * 60 - 1e-6) / 시각분) if 시각_올림 else math.floor((원시작 * 60 + 시각분 / 2) / 시각분)) * 시각분 / 60
                    if not 주말:
                        v = max(v, 평일_가장이른시각)
                else:
                    v = 토요일_시작시각 if 주말 else 시작시각
                시작들.append((v, v + 아동시간[a][d]))
                아동일[a][d]['계획시작'] = v
            처음 = min(x[0] for x in 시작들)
            if 공통_시작_고정:                  # 같은 반 시간표는 항상 기본 시작 시각부터 (더 일찍 오는 아동이 있으면 그 시각부터)
                처음 = min(처음, 토요일_시작시각 if 주말 else 시작시각)
            끝들 = [x[1] for x in 시작들]
            주의 = []
            if 주말:
                슬롯 = 토요일_공통(d, max(끝들) - 처음, tt, 주의, 시작=처음)
                시작 = 처음
                if not 토요일_통째외부 and min(시간들) < 토요일_외부끝 - 토요일_시작시각:
                    주의.append((d, f'토요일 결제 시간이 {시간글자(min(시간들))}시간으로 짧아 10~15시 외부활동을 다 못 채웠어요'))
            else:
                시작 = 처음
                슬롯 = 평일_공통(d, 시작, min(끝들) - 처음, max(끝들) - 처음, tt, 주의)
            날정보[d] = dict(slots=슬롯, start=시작, 바꿈수=0)
            날정보[d]['주활동'] = [] if d.weekday() >= 5 else [m['이름'] for m in 오늘의_주활동(d, tt, [])]
            for dd, msg in 주의:
                점검.append((f'({선생님} 선생님 반)', dd, msg))
        그룹들.append((선생님, 모둠, 날정보))
        for a in 모둠:
            내날정보[a] = 날정보

    # ---- 몇 명만 하는 활동 (예: 카엘댄스 4명): 그 날 그 활동이 있는 아동 중 랜덤으로 N명만
    #      '마트/카엘댄스' 처럼 나란히 적힌 칸이면 뽑힌 아동은 카엘댄스, 나머지는 마트 / 혼자 적힌 칸이면 뽑히지 않은 아동은 대신 활동
    제외맵, 선택맵 = {}, {}
    for 글, 인원 in 일부만_하는_활동.items():
        뽑힌적 = set()
        for d in sorted({d for _, _, 정 in 그룹들 for d in 정}):
            후보 = []
            for 선생님, 모둠, 정 in 그룹들:
                if d not in 정:
                    continue
                for a in 모둠:
                    if d not in 아동시간[a]:
                        continue
                    잘 = 자르기(정[d]['slots'], 아동일[a][d].get('계획시작', 정[d]['start']), 아동시간[a][d])
                    if any((x.get('alt') and 글 in x['alt']['text']) or (not x.get('alt') and 글 in x['text']) for x in 잘):
                        후보.append(a)
            if not 후보 or len(후보) <= 인원 and not any(x.get('alt') for _, _, 정 in 그룹들 if d in 정 for x in 정[d]['slots']):
                continue
            rnd = random.Random(f'{연도}-{월}-{글}-{d}')
            우선 = sorted(a for a in 후보 if a not in 뽑힌적)
            나머지 = sorted(a for a in 후보 if a in 뽑힌적)
            rnd.shuffle(우선); rnd.shuffle(나머지)
            뽑힘 = (우선 + 나머지)[:인원]
            뽑힌적.update(뽑힘)
            대안칸 = any(x.get('alt') for _, _, 정 in 그룹들 if d in 정 for x in 정[d]['slots'])
            for a in 후보:
                if a in 뽑힘:
                    선택맵[(a, d)] = True
                elif not 대안칸:
                    제외맵[(a, d)] = 글
            점검.append((f'({글})', d, f'"{글}" 활동은 {len(후보)}명 중 랜덤 {인원}명({", ".join(sorted(뽑힘))})만 넣고, 나머지는 ' + ('원래 칸 활동(예: 마트)대로 했어요' if 대안칸 else f'{일부만_대신_활동 or "자율활동"}으로 바꿨어요')))
    합친제외 = dict(제외맵)
    for (a, d) in 선택맵:
        합친제외[('선택', a, d)] = True
    for 선생님, 모둠, 날정보 in 그룹들:
        # ---- 3단계: 모든 아동이 외부+협력 30%를 넘도록 산책으로 바꾸기 → 산책 장소 정하기
        결과비교.update(외부_맞추기(날정보, {a: 아동시간[a] for a in 모둠}, 기준맵,
                                {a: {d: v['계획시작'] for d, v in 아동일[a].items() if '계획시작' in v} for a in 모둠},
                                제외={k: v for k, v in 합친제외.items() if (k[1] if k[0] == '선택' else k[0]) in 모둠}))
        산책_장소_정하기(날정보)

    요약, 요구빈칸 = [], []
    전체 = 엑셀열기(양식경로)                  # 모든 아동을 이 엑셀 파일 하나에 시트로 모읍니다
    원본양식 = 전체.worksheets[0]
    원본양식.title = '원본양식'
    for 아동 in 대상:
        담당 = 담당표[아동]
        날별 = 아동일[아동]
        총, 외부합, 목표, 기준가능 = 결과비교[아동]
        기준 = 기준맵[아동]
        if not 기준가능:
            점검.append((아동, None, f'총 활동시간({시간글자(총)}시간)이 짧아 수급시간 {기준}시간의 30%는 채울 수 없어요. 총 활동시간 기준으로만 맞췄어요'))
        날들, 내부합, 외부합2, 협력합, 송영합 = [], 0, 0, 0, 0
        구분 = '\n' if 활동_구분 == '\n' else ' / '
        for 날짜 in sorted(날별):
            d = 날별[날짜]
            일 = dict(날짜=날짜, 인력=d['인력'] or 담당, 인정=d['인정'], 활동='', 내부=0, 외부=0, 협력=0, 송영=0)
            송 = []
            if 송영_따로_처리 and d.get('송영', 0) > 0:      # 송영: 결제 시각을 30분 단위로 맞춘 시각에서 끝나는 칸
                끝시각 = (math.floor((d['송영시각'] * 60 + 시각_반올림_분 / 2) / 시각_반올림_분) * 시각_반올림_분 / 60
                       if d.get('송영시각') is not None else None)
                if 끝시각 is not None:
                    송 = [dict(start=끝시각 - d['송영'], dur=d['송영'], text='송영서비스', ext=False, coop=False, tag='송영')]
                    일['송영'] = d['송영']
                    송영합 += d['송영']
            잘림 = []
            if d['시간'] > 0:
                잘림 = 자르기(내날정보[아동][날짜]['slots'], d.get('계획시작', 내날정보[아동][날짜]['start']), d['시간'])
                if 합친제외.get(('선택', 아동, 날짜)):          # 뽑힌 아동: '마트/카엘댄스' 칸에서 카엘댄스를 함
                    잘림 = [dict(x, **x['alt']) if x.get('alt') else x for x in 잘림]
                개인자율 = _고르기(자율활동_지정[아동], 날짜) if 자율활동_지정.get(아동) else None
                if 개인자율:                             # 이 아동만의 자율활동
                    잘림 = [dict(x, text=개인자율) if x.get('tag') == '자율' else x for x in 잘림]
                글 = 합친제외.get((아동, 날짜))
                if 글:                                   # 이 아동은 그 날 '몇 명만 하는 활동'에 뽑히지 않음 → 대신 활동
                    대신 = 일부만_대신_활동 or 개인자율 or 자율(날짜)['text']
                    잘림 = [dict(x, text=대신, ext=False, coop=False, tag='자율', conv=False) if 글 in x['text'] else x for x in 잘림]
                일['내부'], 일['외부'], 일['협력'] = 합계(잘림)
                내부합 += 일['내부']
                외부합2 += 일['외부'] + 일['협력']
                협력합 += 일['협력']
            if not 송영_글에_쓰기:
                송_글 = []
            else:
                송_글 = 송
            if 잘림 or 송_글:
                일['활동'] = 슬롯_글(sorted(잘림 + 송_글, key=lambda z: (z['start'], z.get('tag') != '송영')))
            대체 = 대체인력_문구.replace('{대체}', 일['인력']) if (d['시간'] > 0 and 일['인력'] != 담당) else ''
            인정글 = (인정결제_문구.replace('{n}', 시간글자(일["인정"])) + (('\n' + 인정결제_덧붙임) if 인정결제_덧붙임 else '')) if (일['인정'] and 인정결제_표기) else ''
            소급 = (' / '.join(소급_문구.get(x, f'소급결제 사유: {x}') for x in d['사유'])) if d['사유'] and 소급_문구 else \
                   (('소급결제 사유: ' + ', '.join(d['사유'])) if d['사유'] else '')
            특이인정 = '' if 인정결제_요구사항에 else 인정글
            일['특이_계획'] = 구분.join(x for x in (('' if 대체_기록지에만 else 대체),) if x)
            일['특이_기록'] = 구분.join(x for x in (대체, 소급, 특이인정) if x)
            날들.append(일)

        비율기준 = 외부합2 / 기준
        비율총 = 외부합2 / 총 if 총 else 0
        if 외부합2 <= 외부_최소비율 * 총 + 1e-9:
            점검.append((아동, None, f'외부+협력이 총시간의 {비율총:.1%}라서 30%를 못 넘었어요 (바꿀 보조활동 부족)'))
        elif 기준가능 and 비율기준 <= 외부_최소비율 + 1e-9:
            점검.append((아동, None, f'외부+협력이 수급시간 {기준}시간의 {비율기준:.1%}라서 30%를 못 넘었어요 (바꿀 보조활동 부족)'))

        요일들 = defaultdict(list)
        for 일 in 날들:
            if 일['활동']:
                요일들[일['날짜'].weekday()] += 내날정보[아동][일['날짜']].get('주활동', [])
        산책있음 = any(s['tag'] == '산책' for d in 아동시간[아동] for s in
                   자르기(내날정보[아동][d]['slots'], 날별[d].get('계획시작', 내날정보[아동][d]['start']), 아동시간[아동][d]))
        인정날 = [일['날짜'] for 일 in 날들 if 일['인정']]
        계획일 = 전달_마지막_근무일(연도, 월)
        if 기록일_방식 == '말일':
            기록일 = 이번달_마지막_근무일(연도, 월)
            요구날짜 = max(인정날) if 인정날 else None
        elif 인정날:
            기록일 = 요구날짜 = max(인정날)
        else:
            기록일 = 이번달_마지막_근무일(연도, 월)
            요구날짜 = 날들[-1]['날짜'] if 날들 else None
            점검.append((아동, None, f'인정결제가 없어서 기록지 작성일자를 {한줄(기록일)}로 넣었어요. 구청에 신청한 날짜가 따로 있으면 알려주세요'))
        요구글 = 요구사항_글(아동, 요일들, 요구목록, 산책있음)
        if not 요구사항_자동:
            요구글 = ''
        elif 요구사항_맨아래 or 요구사항_특이사항:     # 한 줄 방식: 머리말과 '- ' 는 빼고 내용만
            요구글 = ' / '.join(x.lstrip('- ').strip() for x in 요구글.split('\n')[1:] if x.strip())
        if 요구사항_특이사항 and 날들 and 요구글:       # 기록지 마지막 날짜 줄의 개별 특이사항에 '요구사항: ...'
            끝날 = 날들[-1]
            끝날['특이_기록'] = 구분.join(x for x in (끝날['특이_기록'], '요구사항: ' + 요구글) if x)
        총인정 = sum(일['인정'] for 일 in 날들)
        표시날들 = [일 for 일 in 날들 if 일['활동'] or 일['특이_계획'] or 일['특이_기록']] if 인정결제만_있는_날_줄_빼기 else 날들
        정보 = dict(아동=아동, 월=월, 담당=담당, 날들=표시날들, 총=총 + 송영합, 기준=기준, 계획일=계획일, 기록일=기록일,
                  대체인력=sorted({일['인력'] for 일 in 날들 if 일['활동'] and 일['인력'] != 담당}),
                  요구사항=요구글, 요구_줄날짜=요구날짜,
                  인정글=(인정결제_문구.replace('{n}', 시간글자(총인정)) if (인정결제_요구사항에 and 총인정 and 인정결제_표기) else ''),
                  개인예산=(아동 in 수급시간_바꾸기))
        if 요구사항_자동 and (요구사항_맨아래 or 요구사항_특이사항) and not 요구글:
            요구빈칸.append(아동)

        계획 = 전체.copy_worksheet(원본양식)
        계획.title = f'{아동}{시트이름_구분}계획서'
        기록 = 전체.copy_worksheet(원본양식)
        기록.title = f'{아동}{시트이름_구분}기록지'
        시트_채우기(계획, 정보, 계획서=True)
        시트_채우기(기록, 정보, 계획서=False)

        충족 = 외부합2 > 외부_최소비율 * max(기준 if 기준가능 else 0, 총) + 1e-9
        요약.append((아동, 담당, len([d for d in 날들 if d['활동']]), len(인정날), 총, 내부합, 외부합2 - 협력합, 협력합,
                   비율기준, 비율총, '예' if 충족 else '아니오', 한줄(기록일)))
        print(f'  작성 완료: {아동} (담당 {담당})')

    if 요구빈칸:
        점검.append(('(전체)', None, f'"이용자 요구사항)" 칸이 비어 있어요 ({len(요구빈칸)}명). 시간표 파일에 "요구사항" 시트가 없어서 아이마다 직접 적어 주셔야 해요'))
    for ws in list(전체.worksheets):
        if ws.title == '원본양식' or not (ws.title.endswith(시트이름_구분 + '계획서') or ws.title.endswith(시트이름_구분 + '기록지')):
            전체.remove(ws)
    전체파일 = os.path.join(출력, f'방과후활동_{월}월_계획서_기록지.xlsx')
    전체.save(전체파일)
    print(f'\n  → {폴더이름}/{os.path.basename(전체파일)}  (아동 {len(대상)}명, 시트 {len(전체.worksheets)}개)')

    # 점검표
    wb = Workbook()
    s1 = wb.active
    s1.title = '요약'
    s1.append(['아동', '담당 제공인력', '서비스일수', '인정결제 일수', '총 계획시간', '내부(시간)', '외부(시간)', '협력(시간)',
               '외부+협력 ÷ 수급시간', '외부+협력 ÷ 총 계획시간', '30% 초과', '기록지 작성일자'])
    for r in 요약:
        s1.append(list(r))
    for row in s1.iter_rows(min_row=2, min_col=9, max_col=10):
        for c in row:
            c.number_format = '0.0%'
    s2 = wb.create_sheet('확인할 점')
    s2.append(['아동', '날짜', '내용'])
    본 = set()
    for 아동, d, msg in 점검:
        키 = (아동, d, msg)
        if 키 in 본:
            continue
        본.add(키)
        s2.append([아동, 한줄(d) if d else '(전체)', msg])
    for ws in (s1, s2):
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill('solid', fgColor='DDEBF7')
            c.alignment = Alignment(wrap_text=True, vertical='center')
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 16
    s2.column_dimensions['C'].width = 80
    wb.save(os.path.join(출력, '점검표.xlsx'))
    print(f'\n끝! 아동 {len(대상)}명 · 확인할 점 {len(본)}건 → "{폴더이름}" 폴더(바탕화면)의 점검표.xlsx 를 먼저 열어보세요.')
    try:
        if os.name == 'nt' and not 화면에서_실행:
            os.startfile(출력)
    except Exception:
        pass
    return dict(출력=출력, 폴더이름=폴더이름, 월=월, 아동수=len(대상), 확인할점=len(본), 전체파일=전체파일,
                점검표=os.path.join(출력, '점검표.xlsx'),
                요약=[dict(아동=r[0], 담당=r[1], 서비스일수=r[2], 인정일수=r[3], 총시간=r[4], 외부시간=r[6], 통과=(r[10] == '예')) for r in 요약])


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
    if getattr(sys, 'frozen', False):
        input('\n엔터 키를 누르면 창이 닫혀요...')
