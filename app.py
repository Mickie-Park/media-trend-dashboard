import os
import glob
import re
import pandas as pd
import streamlit as st
import plotly.express as px
import google.generativeai as genai

# 1. 페이지 설정
st.set_page_config(page_title="월간 업계 동향 통합 인텔리전스", layout="wide", page_icon="📊")

DATA_DIR = "./data"
os.makedirs(DATA_DIR, exist_ok=True)

# 엑셀 종합 파서
@st.cache_data
def load_all_data():
    all_files = glob.glob("업계동향_*.xlsx") + glob.glob(f"{DATA_DIR}/*.xlsx")
    all_files = sorted(list(set(all_files)))
    
    issues_list = []
    tv_sales_list = []
    pt_list = []
    agency_sales_list = []
    
    for f in all_files:
        m = re.search(r'(\d{4})(\d{2})', f)
        ym = f"{m.group(1)}-{m.group(2)}" if m else "기타"
        
        try:
            xls = pd.ExcelFile(f)
            df = pd.read_excel(f, sheet_name=xls.sheet_names[0])
            
            # A. 주요 Issue 요약
            current_issue = ""
            for idx, val in df['Unnamed: 0'].dropna().items():
                text = str(val).strip()
                if "광고회사 PT" in text:
                    break
                if text.startswith("★"):
                    current_issue = text.replace("★", "").strip()
                    issues_list.append({"연월": ym, "헤드라인": current_issue, "상세": ""})
                elif current_issue and text.startswith("-"):
                    issues_list.append({"연월": ym, "헤드라인": current_issue, "상세": text[1:].strip()})
            
            # B. 광고회사 PT 현황
            pt_start_row, pt_end_row = -1, -1
            for i in range(len(df)):
                row_str = " ".join([str(x) for x in df.iloc[i].dropna().tolist()])
                if "광고회사 PT" in row_str and "종합편" in row_str:
                    pt_start_row = i + 2
                elif "광고회사 전파광고" in row_str and pt_start_row != -1:
                    pt_end_row = i
                    break
            
            if pt_start_row != -1 and pt_end_row != -1:
                for r in range(pt_start_row, pt_end_row):
                    row_vals = df.iloc[r].tolist()
                    pt_date = row_vals[0]
                    client = row_vals[1]
                    product = row_vals[2]
                    billing = row_vals[3]
                    participants = row_vals[4] if len(row_vals) > 4 else ""
                    incumbent = row_vals[5] if len(row_vals) > 5 else ""
                    winner = row_vals[6] if len(row_vals) > 6 else ""
                    memo = row_vals[7] if len(row_vals) > 7 else ""
                    
                    if pd.notna(client) and str(client).strip() != "광고주":
                        date_str = pt_date.strftime("%Y-%m-%d") if isinstance(pt_date, pd.Timestamp) else (str(pt_date).strip() if pd.notna(pt_date) else "")
                        billing_val = 0.0
                        b_match = re.search(r'(\d+)', str(billing))
                        if b_match:
                            try: billing_val = float(b_match.group(1))
                            except: pass
                            
                        pt_list.append({
                            "발행연월": ym,
                            "PT일자": date_str,
                            "광고주": str(client).strip(),
                            "품목": str(product).strip() if pd.notna(product) else "",
                            "빌링(억원)": billing_val,
                            "빌링_원문": str(billing).strip() if pd.notna(billing) else "",
                            "참여사": str(participants).strip() if pd.notna(participants) else "",
                            "기존사": str(incumbent).strip() if pd.notna(incumbent) else "",
                            "선정사": str(winner).strip() if pd.notna(winner) else "",
                            "메모": str(memo).strip() if pd.notna(memo) else ""
                        })

            # C. 대행사 전파광고 매출
            for i in range(len(df)):
                row_str = " ".join([str(x) for x in df.iloc[i].dropna().tolist()])
                if "광고회사 전파광고" in row_str:
                    for offset in range(3, 40):
                        if i + offset >= len(df): break
                        agency_row = df.iloc[i + offset].tolist()
                        name_candidate = str(agency_row[0]).strip() if pd.notna(agency_row[0]) else ""
                        if any(ag in name_candidate for ag in ["제일기획", "이노션", "대홍기획", "HS AD", "TBWA", "SM C&C", "Dentsu"]):
                            try:
                                val = float(str(agency_row[1]).replace(',', '').strip())
                                agency_sales_list.append({"연월": ym, "대행사": name_candidate, "매출(억원)": val})
                            except: pass

            # D. 방송사 매출
            for i in range(len(df)):
                row_vals = [str(x) for x in df.iloc[i].dropna().tolist()]
                if any("지상파 매출" in x for x in row_vals):
                    for offset in range(1, 10):
                        if i + offset >= len(df): break
                        sub_row = df.iloc[i+offset].dropna().tolist()
                        if len(sub_row) >= 3:
                            ch_name = str(sub_row[0]).strip()
                            if any(tv in ch_name for tv in ["KBS", "MBC", "SBS", "Total"]):
                                try:
                                    val = float(str(sub_row[1]).replace(',', '').strip())
                                    tv_sales_list.append({"연월": ym, "채널": ch_name, "매출(억원)": val, "구분": "지상파"})
                                except: pass

                if any("종합/유선채널" in x for x in row_vals):
                    for offset in range(1, 15):
                        if i + offset >= len(df): break
                        sub_row = df.iloc[i+offset].dropna().tolist()
                        if len(sub_row) >= 3:
                            ch_name = str(sub_row[0]).strip()
                            if any(tv in ch_name for tv in ["TV Chosun", "MBN", "Channel A", "JTBC", "SBS Plus", "SPOTV"]):
                                try:
                                    val = float(str(sub_row[1]).replace(',', '').strip())
                                    tv_sales_list.append({"연월": ym, "채널": ch_name, "매출(억원)": val, "구분": "종편/유선"})
                                except: pass
        except Exception as e:
            st.error(f"{f} 파싱 오류: {e}")
            
    return pd.DataFrame(issues_list), pd.DataFrame(tv_sales_list), pd.DataFrame(pt_list), pd.DataFrame(agency_sales_list), all_files

df_issues, df_tv, df_pt, df_agency, loaded_files = load_all_data()

df_pt_unique = df_pt.drop_duplicates(subset=["PT일자", "광고주", "품목"]) if not df_pt.empty else pd.DataFrame()

# --- 사이드바: 관리자 및 AI 키 설정 ---
st.sidebar.title("⚙️ 설정 및 제어판")

# AI API Key 입력 필드
api_key = st.sidebar.text_input("🔑 Gemini API Key (선택)", type="password", help="Google AI Studio에서 발급받은 키를 입력하면 AI 탭이 활성화됩니다.")

role = st.sidebar.radio("접속 권한", ["팀원 (조회 전용)", "마스터 (파일 관리)"])

if role == "마스터 (파일 관리)":
    pwd = st.sidebar.text_input("마스터 비밀번호", type="password")
    if pwd == "admin1234":
        st.sidebar.success("인증 완료: 새 월간 파일을 추가할 수 있습니다.")
        new_file = st.sidebar.file_uploader("월간 엑셀 추가 (.xlsx)", type=["xlsx"])
        if new_file is not None:
            save_path = os.path.join(DATA_DIR, new_file.name)
            with open(save_path, "wb") as f:
                f.write(new_file.getbuffer())
            st.sidebar.success(f"{new_file.name} 저장 완료!")
            st.cache_data.clear()
            st.rerun()
    elif pwd:
        st.sidebar.error("비밀번호가 올바르지 않습니다.")
else:
    st.sidebar.info("💡 팀원 열람 모드입니다.")

st.sidebar.markdown("---")
st.sidebar.write(f"📁 적재 완료 파일: **{len(loaded_files)}건**")

# --- 메인 대시보드 ---
st.title("📊 월간 미디어·광고 업계 동향 대시보드")
st.caption("5년간 축적된 월간 동향 보고서를 다각도로 분석·조회하는 통합 인텔리전스 시스템")

tab1, tab2, tab3, tab4 = st.tabs([
    "🎯 광고회사 PT 수주 현황", 
    "🏢 대행사/매체사 매출 동향", 
    "📰 월별 핵심 이슈 브리핑", 
    "🤖 AI 동향 분석가"
])

# 탭 1: PT 수주 현황
with tab1:
    st.subheader("🎯 광고회사 경쟁 PT 모니터링 & 수주 분석")
    if not df_pt_unique.empty:
        m1, m2, m3 = st.columns(3)
        m1.metric("총 모니터링 PT 건수", f"{len(df_pt_unique)}건")
        m2.metric("집계된 총 빌링 규모", f"{int(df_pt_unique['빌링(억원)'].sum()):,}억원")
        avg_b = df_pt_unique[df_pt_unique['빌링(억원)'] > 0]['빌링(억원)'].mean()
        m3.metric("평균 프로젝트 빌링", f"{avg_b:.1f}억원" if pd.notna(avg_b) else "집계중")

        st.markdown("---")
        f_col1, f_col2, f_col3 = st.columns([2, 2, 2])
        with f_col1:
            search_query = st.text_input("🔍 광고주 또는 품목 검색", placeholder="예: 라이나, 카카오, 샴푸")
        with f_col2:
            min_b = st.slider("최소 빌링 (억원)", 0, int(df_pt_unique['빌링(억원)'].max() if df_pt_unique['빌링(억원)'].max() > 0 else 100), 0)
        with f_col3:
            winner_search = st.text_input("🏆 선정사(승자) 검색", placeholder="예: 제일, 이노션, 차이")

        view_pt = df_pt_unique.copy()
        if search_query:
            view_pt = view_pt[view_pt["광고주"].str.contains(search_query, na=False) | view_pt["품목"].str.contains(search_query, na=False)]
        if min_b > 0:
            view_pt = view_pt[view_pt["빌링(억원)"] >= min_b]
        if winner_search:
            view_pt = view_pt[view_pt["선정사"].str.contains(winner_search, na=False)]

        st.dataframe(
            view_pt[["PT일자", "광고주", "품목", "빌링_원문", "참여사", "기존사", "선정사", "메모"]].rename(columns={"빌링_원문": "빌링(억원)"}),
            width="stretch",
            hide_index=True
        )
    else:
        st.warning("PT 데이터가 아직 없습니다.")

# 탭 2: 매출 동향
with tab2:
    st.subheader("🏢 광고대행사 및 방송 매체사 매출 추이")
    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown("#### 🏆 주요 광고대행사 전파광고 매출")
        if not df_agency.empty:
            fig_agency = px.bar(df_agency, x="대행사", y="매출(억원)", color="연월", barmode="group", text_auto=True)
            st.plotly_chart(fig_agency, use_container_width=True)
        else:
            st.info("대행사 매출 집계 중")
    with col_r:
        st.markdown("#### 📺 방송 매체사 광고 매출")
        if not df_tv.empty:
            fig_tv = px.line(df_tv[df_tv["채널"] != "Total (광고매출 only)"], x="연월", y="매출(억원)", color="채널", markers=True)
            st.plotly_chart(fig_tv, use_container_width=True)
        else:
            st.info("방송사 매출 집계 중")

# 탭 3: 이슈 브리핑
with tab3:
    st.subheader("📰 월별 업계 이슈 & 정책 동향")
    if not df_issues.empty:
        selected_ym = st.selectbox("조회 연월 선택", options=sorted(df_issues["연월"].unique(), reverse=True))
        ym_issues = df_issues[df_issues["연월"] == selected_ym]
        for hl in ym_issues["헤드라인"].unique():
            with st.expander(f"📌 {hl}", expanded=True):
                details = ym_issues[(ym_issues["헤드라인"] == hl) & (ym_issues["상세"] != "")]["상세"].tolist()
                for d in details:
                    st.write(f"- {d}")
    else:
        st.warning("이슈 데이터가 없습니다.")

# 탭 4: AI 동향 분석가
with tab4:
    st.subheader("🤖 AI 기반 업계 동향 분석가")
    if not api_key:
        st.warning("👈 왼쪽 사이드바에 'Gemini API Key'를 입력하시면 실시간 질의응답이 활성화됩니다.")
    else:
        try:
            genai.configure(api_key=api_key)
            
            # 최신 기본 모델(gemini-3.6-flash) 지정
            target_model = "gemini-3.6-flash"
            model = genai.GenerativeModel(target_model)
            st.caption(f"연결된 AI 모델: `{target_model}`")
            
            user_question = st.text_input("질문을 입력하세요", placeholder="예: 디즈니플러스 런칭이 유료방송에 미친 영향은?")
            
            if st.button("AI 분석 요청", type="primary") and user_question:
                with st.spinner("동향 데이터를 분석 중입니다..."):
                    # 컨텍스트 최적화
                    context_issues = df_issues.head(40).to_string(index=False)
                    context_pt = df_pt_unique.head(30).to_string(index=False)
                    
                    prompt = f"""
당신은 대한민국 미디어·방송·광고 업계 전문 분석가입니다.
아래 제공된 [월별 업계 이슈 데이터]와 [광고회사 PT 현황 데이터]를 기반으로 질문에 명확하고 간결하게 답변해 주세요.

[월별 업계 이슈 데이터]
{context_issues}

[최근 주요 PT 현황]
{context_pt}

[질문]
{user_question}

지침:
1. 제공된 데이터의 구체적 사실(기업명, 수치 등)을 기반으로 작성하세요.
2. 가독성을 위해 불릿 포인트로 정리하세요.
3. 데이터에 없는 내용은 추측하지 마세요.
"""
                    response = model.generate_content(prompt)
                    st.markdown("### 💡 AI 분석 리포트")
                    st.markdown(response.text)
        except Exception as e:
            err_msg = str(e)
            if "429" in err_msg:
                st.error("⚠️ 일시적으로 무료 호출 한도에 도달했습니다. 잠시 후 다시 시도해 주세요.")
            else:
                st.error(f"AI 호출 오류: {err_msg}")