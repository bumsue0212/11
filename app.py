import streamlit as st
import pandas as pd
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO
import matplotlib.pyplot as plt
from matplotlib import rc

# 한글 폰트 설정
rc('font', family='Malgun Gothic')  # Windows 사용자는 'Malgun Gothic', Mac 사용자는 'AppleGothic'

# VLOOKUP 구현
def vlookup(value, lookup_table, key_col, value_col):
    try:
        return lookup_table.loc[lookup_table[key_col] == value, value_col].values[0]
    except IndexError:
        return 0

# "수기관리" 기본 파일 생성
def create_default_reference_file(reference_file_path):
    if not os.path.exists(reference_file_path):
        with pd.ExcelWriter(reference_file_path, engine="openpyxl") as writer:
            # 박스 정보 시트
            box_info = pd.DataFrame({
                "박스번호": ["○단독박스", "단독박스", "B-42", "B-66", "B-56", "C-26", "B-92", "C-140", "B-149", "B-153", "B-160", "B-170"],
                "박스가격(VAT포함)": [0, 0, 171, 176, 220, 270, 312, 438, 683, 1098, 604, 520],
            })
            box_info.to_excel(writer, sheet_name="박스 정보", index=False)

            # 기본 비용 시트
            cost_info = pd.DataFrame({
                "항목": ["작업비", "운반비"],
                "금액": [550, 0],
            })
            cost_info.to_excel(writer, sheet_name="기본비용", index=False)

            st.success(f"Default 수기관리 file created at: {reference_file_path}")

# 사용박스 정보 추출 함수
def extract_box_info(product_name, box_info):
    for box_number in box_info["박스번호"].values:
        if box_number in product_name:
            return box_number
    return None

# 데이터 처리 및 저장
def process_file(original_df, reference_box, reference_cost):
    processed_df = pd.DataFrame()

    # 필요한 열 추가
    processed_df["소포주문번호"] = original_df["소포주문번호"].astype(str)
    processed_df["상품명"] = original_df["상품명"]
    processed_df["등기번호"] = original_df["등기번호"].astype(str)
    processed_df["접수일시"] = original_df["접수일시"]
    processed_df["상품주문번호"] = original_df["상품주문번호"]
    processed_df["요금"] = original_df["요금"]
    processed_df["공급지"] = original_df["공급지"]

    # 사용박스 정보 추출
    processed_df["사용박스"] = original_df["상품명"].apply(lambda x: extract_box_info(str(x), reference_box))

    # 식별 가능한 박스 가격 계산
    processed_df["식별가능박스가격"] = processed_df["사용박스"].apply(
        lambda x: vlookup(x, reference_box, "박스번호", "박스가격(VAT포함)")
    )

    # 식별 불가능 박스 처리
    processed_df["식별불가능박스가격일괄300원적용"] = processed_df["사용박스"].apply(
        lambda x: 0 if x in reference_box["박스번호"].values else 300
    )

    # 작업비 및 운반비 추가
    processed_df["(3PL)작업비"] = reference_cost.loc[reference_cost["항목"] == "작업비", "금액"].values[0]
    processed_df["(의정부집중국)운반비"] = reference_cost.loc[reference_cost["항목"] == "운반비", "금액"].values[0]

    # 총 택배비 계산
    processed_df["(우체국택배)택배비+부자재+작업비+운반비등"] = (
        processed_df["식별가능박스가격"]
        + processed_df["식별불가능박스가격일괄300원적용"]
        + processed_df["요금"]
        + processed_df["(3PL)작업비"]
        + processed_df["(의정부집중국)운반비"]
    )

    # 날짜형식 처리
    processed_df["접수일시(날짜형식)"] = pd.to_datetime(processed_df["접수일시"], errors="coerce")
    processed_df["연도(접수일)"] = processed_df["접수일시(날짜형식)"].dt.year
    processed_df["월(접수일)"] = processed_df["접수일시(날짜형식)"].dt.month
    processed_df["일(접수일)"] = processed_df["접수일시(날짜형식)"].dt.day

    return processed_df

# 도넛 그래프 생성 함수
def create_donut_chart(data, labels, title):
    colors = plt.cm.tab10.colors[:len(data)]
    explode = [0.1 if i == 0 else 0 for i in range(len(data))]  # 첫 번째 조각 강조

    fig, ax = plt.subplots()
    ax.pie(data, explode=explode, labels=labels, colors=colors, autopct="%1.1f%%", startangle=140, pctdistance=0.85)
    centre_circle = plt.Circle((0, 0), 0.70, fc="white")
    fig.gca().add_artist(centre_circle)
    ax.axis('equal')
    plt.title(title)
    return fig

# Streamlit UI
def main():
    st.title("우체국 택배 데이터 처리")

    # 원본 파일 선택
    original_file = st.file_uploader("원본 파일 업로드 (Excel)", type=["xlsm", "xlsx"])
    reference_file_path = "수기관리.xlsx"

    # 기본 수기관리 파일 생성
    create_default_reference_file(reference_file_path)

    reference_box = pd.read_excel(reference_file_path, sheet_name="박스 정보")
    reference_cost = pd.read_excel(reference_file_path, sheet_name="기본비용")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("박스 정보 수정")
        reference_box = st.data_editor(reference_box, key="박스 정보 수정")

    with col2:
        st.subheader("기본 비용 수정")
        reference_cost = st.data_editor(reference_cost, key="기본 비용 수정")

    if original_file:
        original_df = pd.read_excel(original_file, header=5)

        # 처리된 데이터 생성
        processed_df = process_file(original_df, reference_box, reference_cost)

        # 처리 결과 표시
        st.dataframe(processed_df)

        # 추가 정보 계산
        total_cost = processed_df["(우체국택배)택배비+부자재+작업비+운반비등"].sum()
        box_usage = processed_df["사용박스"].value_counts()
        unidentified_count = processed_df[processed_df["사용박스"].isna()].shape[0]
        identified_count = processed_df[~processed_df["사용박스"].isna()].shape[0]
        total_boxes = identified_count + unidentified_count

        start_date = processed_df["접수일시(날짜형식)"].min()
        end_date = processed_df["접수일시(날짜형식)"].max()

        st.metric("총 택배비", f"{total_cost:,} 원")
        st.metric("총 박스 수", f"{total_boxes} 개")
        st.metric("처리 기간", f"{start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}")

        # 도넛 그래프 생성 및 표시
        fig1 = create_donut_chart([identified_count, unidentified_count], ["식별된 박스", "식별되지 않은 박스"], "박스 식별 현황")
        st.pyplot(fig1)

        identified_box_counts = box_usage.values
        identified_box_labels = box_usage.index
        fig2 = create_donut_chart(identified_box_counts, identified_box_labels, "식별된 박스 구성")
        st.pyplot(fig2)

        # 다운로드 버튼 추가
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            processed_df.to_excel(writer, index=False, sheet_name="Processed Data")
            reference_box.to_excel(writer, index=False, sheet_name="박스 정보")
            reference_cost.to_excel(writer, index=False, sheet_name="기본비용")
        output.seek(0)

        st.download_button(
            label="Download Processed File",
            data=output,
            file_name="processed_result_with_reference.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if __name__ == "__main__":
    main()
