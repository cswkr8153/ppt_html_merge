PPT / HTML 포맷 일치화 도구
===========================
요구: Python 3.8+ (표준 라이브러리만 사용, pip 설치 불필요). GUI는 tkinter 포함 Python 필요.

실행 (GUI)
  python format_unifier.py
  1) 출력 형식(PPT/HTML) 선택 → 2) 기준 파일 가져오기 → 3) 대상 파일 추가 → 4) 실행 → 5) ZIP 저장

실행 (CLI)
  python format_unifier.py --to ppt  --base 기준.pptx a.pptx b.html --out 결과폴더
  python format_unifier.py --to html --base 기준.html a.html b.pptx --out 결과폴더

동작
  - 원본은 수정하지 않고 "<이름>_unified.pptx/.html" 새 파일로 저장, 변경요약.txt 포함 ZIP 생성
  - PPT 출력: 기준 파일의 슬라이드 크기/마스터/레이아웃/테마를 유지하고 대상의 내용(텍스트·목록·표·이미지)만 이식
  - HTML 출력: 기준 파일의 head(CSS/메타)·헤더·푸터·레이아웃에 대상 본문을 삽입, 인라인 스타일·구식 태그 정리
  - 기준이 PPT이고 출력이 HTML이면 PPT 테마(폰트/색/비율)로 CSS를 생성
  - 현재 미지원: HTML 기준 → PPT 출력 (PPT 기준 필요)

사내 암호화(DRM) PPT
  Windows + PowerPoint 설치 환경에서 자동으로 COM으로 열어 일반 pptx 복사본을 만들어 처리하고
  임시 파일은 처리 후 삭제합니다. (pywin32가 있으면 사용, 없으면 PowerShell 사용)
  결과 pptx 재암호화는 포함하지 않습니다.

한계
  차트/SmartArt/애니메이션/복잡한 도형, HTML의 JS/복잡한 레이아웃은 변환 시 단순화되거나 누락됩니다(변경요약에 표시).
