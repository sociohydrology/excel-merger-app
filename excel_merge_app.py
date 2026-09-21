import pandas as pd
import streamlit as st
import io

st.set_page_config(page_title="스마트 엑셀 취합기", layout="wide")

st.title("📁 스마트 엑셀 취합 프로그램")
st.markdown("공통 항목은 표준명으로 묶고, 나머지 고유 열은 원본 헤더 그대로 살려 병합합니다.")

# 1. 파일 업로드
uploaded_files = st.file_uploader(
    "병합할 엑셀 파일들을 선택해주세요 (다중 선택 가능)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"총 {len(uploaded_files)}개의 파일이 업로드되었습니다.")
    
    dataframes = []
    file_info = []

    # 각 파일 읽기
    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file)
            df['출처_파일명'] = uploaded_file.name 
            dataframes.append(df)
            file_info.append({
                "파일명": uploaded_file.name,
                "컬럼 목록": list(df.columns)
            })
        except Exception as e:
            st.error(f"파일을 읽는 중 오류 발생 ({uploaded_file.name}): {e}")

    if dataframes:
        st.subheader("📊 업로드된 파일 구조 확인")
        
        with st.expander("파일별 상세 컬럼 보기"):
            for info in file_info:
                st.write(f"**{info['파일명']}**: {info['컬럼 목록']}")

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
                cols = st.columns(len(standard_cols))
                
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

        if st.button("🚀 엑셀 합치기 실행"):
            with st.spinner("데이터를 정제하고 병합하는 중입니다..."):
                try:
                    processed_dfs = []
                    
                    for df, mapping in zip(dataframes, mapping_configs):
                        # 원본 데이터 복사
                        temp_df = df.copy()
                        
                        # 사용자가 매핑한 컬럼들만 이름 변경 (나머지 헤더는 그대로 유지됨)
                        if mapping:
                            temp_df = temp_df.rename(columns=mapping)
                        
                        processed_dfs.append(temp_df)

                    # 데이터 통합 (서로 다른 이름의 컬럼은 자동으로 공간을 넓혀서 병합됨)
                    merged_df = pd.concat(processed_dfs, ignore_index=True, sort=False)

                    st.success("성공적으로 병합되었습니다!")
                    st.dataframe(merged_df.head(10))

                    # 엑셀 다운로드 버튼 생성
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        merged_df.to_excel(writer, index=False, sheet_name='통합데이터')
                    
                    processed_data = output.getvalue()

                    st.download_button(
                        label="📥 통합된 엑셀 파일 다운로드 (.xlsx)",
                        data=processed_data,
                        file_name="스마트_통합보고서.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"병합 중 오류가 발생했습니다: {e}")