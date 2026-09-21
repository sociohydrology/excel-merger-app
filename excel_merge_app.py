import pandas as pd
import streamlit as st
import io

st.set_page_config(page_title="스마트 엑셀 취합기", layout="wide")

st.title("📁 스마트 엑셀 취합 프로그램")
st.markdown("양식이 조금 달라도 걱정마세요! 여러 개의 엑셀 파일을 업로드하여 깔끔하게 하나로 병합해 드립니다.")
st.markdown("업로드한 파일은 서버에 저장되지 않고, 병합 작업 즉시 메모리에서 안전하게 파기됩니다.")

# 1. 파일 업로드
uploaded_files = st.file_uploader(
    "병합할 엑셀 파일들을 선택해주세요 (다중 선택 가능)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"총 {len(uploaded_files)}개의 파일이 업로드되었습니다.")
    
    # [기능 추가 1] 상단 빈 행 자동 스킵을 위한 탐색 행 수 설정 옵션
    with st.expander("⚙️ 고급 설정 (상단 제목/빈 행 자동 감지 옵션)"):
        skip_rows_limit = st.slider(
            "파일 상단에서 헤더(컬럼명)를 찾기 위해 스캔할 최대 행 수",
            min_value=0, max_value=10, value=5,
            help="엑셀 파일 맨 위에 로고나 빈 줄이 있을 때, 몇 번째 줄까지 내려가서 진짜 표의 시작(헤더)을 찾을지 설정합니다."
        )

    dataframes = []
    file_info = []

    # 각 파일 읽기 및 지능형 헤더 감지 로직
    for uploaded_file in uploaded_files:
        try:
            # 먼저 헤더 없이 전체를 읽어옴 (최대 탐색 행 수만큼)
            temp_raw_df = pd.read_excel(uploaded_file, header=None, nrows=skip_rows_limit + 5)
            
            header_row_idx = 0
            # 위에서부터 탐색하며 데이터가 가장 많이 채워져 있는 행을 헤더(컬럼명)로 추정
            for idx, row in temp_raw_df.iterrows():
                if idx > skip_rows_limit:
                    break
                # 비어있지 않은 셀의 개수가 2개 이상이고, 문자열이 포함되어 있으면 헤더로 판단
                non_empty_count = row.dropna().count()
                if non_empty_count >= 2:
                    header_row_idx = idx
                    break
            
            # 찾아낸 진짜 헤더 행 위치를 반영하여 엑셀 다시 읽기
            df = pd.read_excel(uploaded_file, header=header_row_idx)
            
            # 컬럼명에 공백이 있거나 비어있는 경우 정리
            df.columns = [str(col).strip() for col in df.columns]
            # 빈 이름의 컬럼이 생겼다면 임시 이름 부여
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

            df['출처_파일명'] = uploaded_file.name 
            dataframes.append(df)
            file_info.append({
                "파일명": uploaded_file.name,
                "인식된 헤더행": f"{header_row_idx + 1}번째 행",
                "컬럼 목록": list(df.columns)
            })
        except Exception as e:
            st.error(f"파일을 읽는 중 오류 발생 ({uploaded_file.name}): {e}")

    if dataframes:
        st.subheader("📊 업로드된 파일 구조 확인")
        
        with st.expander("파일별 상세 컬럼 및 헤더 감지 결과 보기"):
            for info in file_info:
                st.write(f"**{info['파일명']}** (인식된 헤더: {info['인식된 헤더행']})")
                st.write(f"-> 컬럼: {info['컬럼 목록']}")

        st.markdown("---")
        st.subheader("⚙️ 표준 컬럼 매칭 설정 (선택 사항)")
        
        standard_cols_input = st.text_input(
            "통합하고 싶은 대표(표준) 컬럼명들을 쉼표(,)로 입력해주세요",
            value="날짜, 품명, 수량, 금액"
        )
        standard_cols = [col.strip() for col in standard_cols_input.split(",") if col.strip()]

        st.markdown("#### 파일별 컬럼 매칭 지정")
        st.markdown("매칭하지 않고 **'선택 안 함'**으로 두면, 기존 파일의 헤더 이름 그대로 병합에 포함됩니다.")

        mapping_configs = []
        for i, (uploaded_file, df) in enumerate(zip(uploaded_files, dataframes)):
            with st.container():
                st.markdown(f"**📄 파일: {uploaded_file.name}**")
                file_cols = list(df.columns)
                
                col_mapping = {}
                cols = st.columns(len(standard_cols)) if standard_cols else [st.container()]
                
                for idx, std_col in enumerate(standard_cols):
                    with cols[idx]:
                        default_index = 0
                        for f_col in file_cols:
                            if std_col in str(f_col):
                                default_index = file_cols.index(f_col) + 1
                                break
                        
                        selected_col = st.selectbox(
                            f"-> [{std_col}]로 변경할 열",
                            options=["선택 안 함"] + file_cols,
                            index=default_index,
                            key=f"map_{i}_{std_col}"
                        )
                        if selected_col != "선택 안 함":
                            col_mapping[selected_col] = std_col
                
                mapping_configs.append(col_mapping)

        st.markdown("---")
        # [기능 추가 2] 다운로드 파일명 커스텀 입력 필드
        custom_file_name = st.text_input(
            "💾 다운로드할 결과 파일명을 입력하세요 (확장자 .xlsx 제외)",
            value="스마트_통합보고서"
        )

        if st.button("🚀 엑셀 합치기 실행"):
            with st.spinner("데이터를 정제하고 병합하는 중입니다..."):
                try:
                    processed_dfs = []
                    
                    for df, mapping in zip(dataframes, mapping_configs):
                        temp_df = df.copy()
                        if mapping:
                            temp_df = temp_df.rename(columns=mapping)
                        processed_dfs.append(temp_df)

                    # 데이터 통합 (서로 다른 이름의 컬럼은 자동으로 공간을 넓혀서 병합됨)
                    merged_df = pd.concat(processed_dfs, ignore_index=True, sort=False)

                    st.success("성공적으로 병합되었습니다!")
                    st.dataframe(merged_df.head(10))

                    # 파일 이름 정제 (사용자가 입력한 이름 + .xlsx)
                    final_file_name = f"{custom_file_name.strip() or '통합보고서'}.xlsx"

                    # 엑셀 다운로드 버튼 생성
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        merged_df.to_excel(writer, index=False, sheet_name='통합데이터')
                    
                    processed_data = output.getvalue()

                    st.download_button(
                        label=f"📥 [{final_file_name}] 다운로드하기",
                        data=processed_data,
                        file_name=final_file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"병합 중 오류가 발생했습니다: {e}")
