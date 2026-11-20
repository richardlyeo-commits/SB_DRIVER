"""
삼영보영 기사 사건·전입전출 입력 도구 - 깃허브 자동 백업 스크립트
=================================================================
이 스크립트가 하는 일 (한 번 실행하면):
 1) Claude 아티팩트의 db(공용 저장소)에서 세 가지 목록을 읽어옵니다.
    - event_records : 사건기록
    - move_records  : 전입전출
    - roster        : 노선별 기사 명단
 2) 세 목록을 "backup_data.json" 파일 하나로 저장합니다.
    (이 json 파일은, 이 파이썬 스크립트가 직접 db를 읽을 수 없기 때문에
     Claude가 먼저 mcp__claude-artifacts__db_list 로 읽어서 넣어주는 자리입니다)
 3) backup_data.json을 읽어서 사람이 보기 좋은 엑셀 파일
    (삼영보영_백업.xlsx, 시트 3개)로 바꿔 저장합니다.

사용법: 이 스크립트 혼자서는 못 돌고, 먼저 backup_data.json을 만들어 둔 다음
       "python3 make_backup.py" 로 실행합니다.
"""
import json
import os
from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(HERE, "backup_data.json")
OUT_PATH = os.path.join(HERE, "backup", "삼영보영_백업.xlsx")

# 시트 이름과, 각 시트에 적을 열(컬럼) 순서를 미리 정해 둠
SHEETS = {
    "사건기록":   ["발생일시", "성명", "노선", "차량", "구분", "접수처", "장소", "위반", "금액", "메모"],
    "전입전출":   ["날짜", "성명", "구분", "개인차량", "메모"],
    "기사명부":   ["노선", "성명"],
}
# json 쪽 컬렉션 이름 -> 위 시트 이름 매칭
COLLECTION_TO_SHEET = {
    "event_records": "사건기록",
    "move_records": "전입전출",
    "roster": "기사명부",
}

def main():
    with open(DATA_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)  # { "event_records": [...], "move_records": [...], "roster": [...] }

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)  # 기본으로 생기는 빈 시트 하나 제거

    for 컬렉션, 시트이름 in COLLECTION_TO_SHEET.items():
        rows = data.get(컬렉션, [])
        columns = SHEETS[시트이름]
        ws = wb.create_sheet(시트이름)
        ws.append(columns)  # 맨 윗줄: 제목(헤더) 줄
        for row in rows:
            ws.append([row.get(col, "") for col in columns])
        # 열 너비를 적당히 넓혀서 읽기 편하게 함
        for i, col in enumerate(columns, start=1):
            ws.column_dimensions[chr(64 + i)].width = max(10, len(col) * 2)

    wb.save(OUT_PATH)
    print(f"저장 완료: {OUT_PATH} (사건기록 {len(data.get('event_records', []))}건, "
          f"전입전출 {len(data.get('move_records', []))}건, 기사명부 {len(data.get('roster', []))}명)")

if __name__ == "__main__":
    main()
