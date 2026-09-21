# 📁 스마트 엑셀 취합 프로그램 (Smart Excel Merger)

양식이 조금 달라도 걱정 마세요! 여러 개의 엑셀 파일을 업로드하여 공통 항목은 표준명으로 묶고, 고유 열은 원본 그대로 살려 깔끔하게 병합해 주는 웹 기반 유틸리티 툴입니다.

---

## ✨ 주요 기능
- **다중 파일 병합:** 여러 개의 `.xlsx`, `.xls` 파일을 드래그 앤 드롭으로 한 번에 취합
- **지능형 컬럼 매칭:** 파일마다 열 이름이 달라도 대표 표준 컬럼명으로 쉽게 매핑 가능
- **고유 열 보존:** 매칭하지 않은 나머지 고유 컬럼들도 누락 없이 원본 그대로 확장 병합
- **웹 기반 인터페이스:** 설치 없이 웹 브라우저에서 간편하게 실행 및 결과 파일 다운로드

---

## 🚀 사용 방법 (웹 버전)
1. 배포된 웹사이트 링크 접속 ([여기에 본인의 Streamlit 웹 링크 입력])
2. 병합할 엑셀 파일들을 마우스로 드래그하여 업로드
3. 필요에 따라 표준 컬럼명 설정 및 파일별 열 매칭 지정
4. **[엑셀 합치기 실행]** 버튼 클릭 후 통합된 파일 다운로드

---

## 🛠️ 기술 스택 (Tech Stack)
- **Language:** Python
- **Framework:** Streamlit
- **Data Processing:** Pandas
- **Excel Engine:** Openpyxl, Xlsxwriter

---

## 💻 로컬 실행 방법 (For Developers)
프로젝트를 로컬 환경에서 직접 실행하고 싶다면 아래 명령어를 입력하세요.

```bash
# 1. 저장소 클론 또는 다운로드
git clone [https://github.com/본인아이디/저장소이름.git](https://github.com/본인아이디/저장소이름.git)
cd 저장소이름

# 2. 필수 패키지 설치
pip install -r requirements.txt

# 3. Streamlit 앱 실행
streamlit run excel_merge_app.py
