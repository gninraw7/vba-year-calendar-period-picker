# vba-year-calendar-period-picker
Excel-based Period Picker &amp; Year Calendar in VBA. JSON-backed tasks, overlay rendering, holidays, i18n (K/E), slot layouts, and export-to-sheet.
# Period Picker (Excel VBA)

엑셀에서 기간 선택과 연간 캘린더를 빠르게 다루는 VBA 솔루션입니다.  
Task를 JSON 파일로 저장/불러오고, 달력에 오버레이하며, 공휴일/주말 제외, 다국어(K/E), 슬롯 배치(현재월 앵커), 시트 출력(툴팁→메모) 등을 지원합니다.

## 주요 기능
- 연간 달력 4×3 슬롯 배치 (Normal / Current First / Current Last / Current @ Slot)
- From~To 선택, 비즈니스 데이(주말·공휴일 제외) 카운트
- Task 패널: JSON 저장/불러오기, 카테고리(파일) 관리, 멀티선택, 선택 범위 하이라이트
- 오버레이: Task별 색/툴팁, 공휴일 표시, 선택범위 강조와 공존
- 시트 출력: 화면과 동일한 오버레이 배경 + 툴팁(휴일/Task만) → 셀 메모
- 국제화: 한/영, 월 타이틀/요일 포맷, 주 시작 요일 변경 지원
- 설정/상태: 레지스트리 저장(창 유지, 자동 줄바꿈, 오버레이, 링크-셀 등)


