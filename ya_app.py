import os
import pandas as pd
import streamlit as st

# 페이지 기본 설정
st.set_page_config(
    page_title="야적장 배치도 검색 시스템", page_icon="🏗️", layout="wide"
)

st.title("🏗️ 야적장 배치도 검색 시스템")

# ---------------------------------------------------------
# 1. 사이드바 - 파일 업로드 및 옵션 설정
# ---------------------------------------------------------
st.sidebar.header("📁 데이터 설정")

uploaded_file = st.sidebar.file_uploader(
    "📂 엑셀 파일(.xlsx) 선택 (기본값: data.xlsx)", type=["xlsx"]
)

# ---------------------------------------------------------
# 2. 데이터 불러오기 (우선순위: 업로드 파일 -> data.xlsx)
# ---------------------------------------------------------
df = None

if uploaded_file is not None:
    # 사용자가 웹페이지에서 직접 엑셀 파일을 업로드한 경우
    try:
        df = pd.read_excel(uploaded_file)
        st.sidebar.success("✅ 사용자 업로드 파일 적용 완료!")
    except Exception as e:
        st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
        st.stop()
elif os.path.exists("data.xlsx"):
    # 업로드한 파일이 없지만, GitHub 저장소에 data.xlsx 파일이 존재하는 경우
    try:
        df = pd.read_excel("data.xlsx")
        st.sidebar.info("ℹ️ GitHub 기본 파일(data.xlsx) 로드 완료")
    except Exception as e:
        st.error(f"기본 data.xlsx 파일을 읽는 중 오류가 발생했습니다: {e}")
        st.stop()
else:
    # 둘 다 없는 경우 안내 출력
    st.warning(
        "⚠️ 등록된 엑셀 데이터가 없습니다.\n\n"
        "1. GitHub 저장소에 **`data.xlsx`** 이름으로 기본 엑셀 파일을 업로드하시거나,\n"
        "2. 왼쪽 사이드바에서 엑셀 파일을 업로드해 주세요."
    )
    st.stop()

# ---------------------------------------------------------
# 3. 데이터 확인 및 검색/배치도 표시 로직
# ---------------------------------------------------------
st.subheader("📋 야적장 데이터 현황")

# 검색 키워드 입력
search_keyword = st.text_input("🔍 검색어 입력 (예: 부품명, 구역, 관리번호 등)")

if search_keyword:
    # 전체 컬럼 중 검색어가 포함된 행 필터링 (대소문자 구분 없이)
    mask = df.astype(str).apply(
        lambda x: x.str.contains(search_keyword, case=False, na=False)
    )
    filtered_df = df[mask.any(axis=1)]

    st.write(
        f"**'{search_keyword}'** 검색 결과: 총 **{len(filtered_df)}** 건"
    )
    st.dataframe(filtered_df, use_container_width=True)
else:
    # 검색어가 없을 때는 전체 데이터 표시
    st.dataframe(df, use_container_width=True)

# ---------------------------------------------------------
# (추가) 시각화 / 배치도 표현 부분이 있다면 아래에 작성
# ---------------------------------------------------------
