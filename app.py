import streamlit as st
import pandas as pd
import glob
import plotly.express as px
import plotly.graph_objects as go
import os

# ==========================================
# 1. 페이지 기본 설정 및 디자인 CSS
# ==========================================
st.set_page_config(
    page_title="맞춤형 축구 배팅 분석 대시보드",
    page_icon="⚽",
    layout="wide"
)

st.markdown("""
<style>
    .main { background-color: #0b0f19; }
    .stApp { background-color: #0b0f19; }
    .stMetric {
        background-color: #1e293b !important;
        border: 1px solid #334155 !important;
        padding: 12px !important;
        border-radius: 10px !important;
    }
    .stMetric label { color: #94a3b8 !important; font-size: 0.85rem !important; }
    .stMetric [data-testid="stMetricValue"] { color: #38bdf8 !important; font-weight: bold !important; font-size: 1.4rem !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #1e293b;
        border-radius: 8px;
        padding: 8px 16px;
        color: #94a3b8;
        border: 1px solid #334155;
    }
    .stTabs [aria-selected="true"] {
        background-color: #3b82f6 !important;
        color: #ffffff !important;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 데이터 로드 및 전처리
# ==========================================
@st.cache_data
def load_data():
    files = sorted(glob.glob("*.xlsx"))
    if not files:
        return {}
    
    league_sheets = {
        'EPL (잉글랜드)': 'E0',
        '라리가 (스페인)': 'SP1',
        '세리에 A (이탈리아)': 'I1',
        '분데스리가 (독일)': 'D1'
    }
    
    league_data = {}
    for league_name, sheet_code in league_sheets.items():
        dfs = []
        for f in files:
            try:
                xls = pd.ExcelFile(f)
                if sheet_code in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_code)
                    season = os.path.basename(f).replace('all-euro-data-', '').replace('.xlsx', '')
                    df['Season'] = season
                    dfs.append(df)
            except Exception:
                pass
        if dfs:
            combined_df = pd.concat(dfs, ignore_index=True)
            combined_df['Date'] = pd.to_datetime(combined_df['Date'], dayfirst=True, errors='coerce')
            combined_df = combined_df.dropna(subset=['HomeTeam', 'AwayTeam', 'FTHG', 'FTAG']).sort_values(by='Date', ascending=False)
            league_data[league_name] = combined_df
            
    return league_data

league_dict = load_data()

st.title("⚽ 선택 매치업 배팅 분석 대시보드")
st.caption("실시간 배당 입력 및 직관적인 지표 비교 시스템이 탑재된 대시보드입니다.")

if not league_dict:
    st.error("❌ `D:\\Data` 폴더에 `.xlsx` 데이터 파일이 없습니다.")
    st.stop()

# ==========================================
# 3. 사이드바 - 리그 및 분석 대상 매치업 선택
# ==========================================
st.sidebar.header("⚙️ 매치업 선택")
selected_league = st.sidebar.selectbox("1. 리그 선택", list(league_dict.keys()))

df = league_dict[selected_league]
teams = sorted(list(set(df['HomeTeam'].dropna().unique()).union(set(df['AwayTeam'].dropna().unique()))))

col_s1, col_s2 = st.sidebar.columns(2)
home_team = st.sidebar.selectbox("2. 홈 팀 선택", teams, index=0)
away_team = st.sidebar.selectbox("3. 원정 팀 선택", [t for t in teams if t != home_team], index=min(1, len(teams)-1))

st.sidebar.markdown("---")
st.sidebar.success(f"🎯 **{home_team} (홈) vs {away_team} (원정)**")

# ==========================================
# 4. 4개 페이지 구성
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs([
    "⚔️ Page 1: 맞대결(H2H) & 최근 흐름",
    "💰 Page 2: 배당 직접 입력 & 수율 분석",
    "🟨 Page 3: 매치 심판 성향",
    "📈 Page 4: 홈 vs 원정 직관적 지표 비교"
])

# ------------------------------------------
# Page 1: 맞대결(H2H) & 최근 흐름
# ------------------------------------------
with tab1:
    st.subheader(f"⚔️ {home_team} vs {away_team} 맞대결 전적 & 최근 흐름")
    
    h2h = df[((df['HomeTeam'] == home_team) & (df['AwayTeam'] == away_team)) | 
             ((df['HomeTeam'] == away_team) & (df['AwayTeam'] == home_team))].sort_values(by='Date', ascending=False)
    
    st.markdown("##### 📌 **역대 맞대결 요약 (H2H)**")
    if not h2h.empty:
        home_wins = len(h2h[((h2h['HomeTeam'] == home_team) & (h2h['FTR'] == 'H')) | ((h2h['AwayTeam'] == home_team) & (h2h['FTR'] == 'A'))])
        away_wins = len(h2h[((h2h['HomeTeam'] == away_team) & (h2h['FTR'] == 'H')) | ((h2h['AwayTeam'] == away_team) & (h2h['FTR'] == 'A'))])
        draws = len(h2h[h2h['FTR'] == 'D'])
        
        h2h['TotalGoals'] = h2h['FTHG'] + h2h['FTAG']
        over25 = len(h2h[h2h['TotalGoals'] > 2.5])
        btts = len(h2h[(h2h['FTHG'] > 0) & (h2h['FTAG'] > 0)])
        
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric(f"{home_team} 승리", f"{home_wins}회", f"{(home_wins/len(h2h))*100:.0f}%")
        m2.metric("무승부", f"{draws}회", f"{(draws/len(h2h))*100:.0f}%")
        m3.metric(f"{away_team} 승리", f"{away_wins}회", f"{(away_wins/len(h2h))*100:.0f}%")
        m4.metric("2.5 오버 비율", f"{over25}경기", f"{(over25/len(h2h))*100:.0f}%")
        m5.metric("양팀 득점 (BTTS)", f"{btts}경기", f"{(btts/len(h2h))*100:.0f}%")
        
        with st.expander("📄 최근 맞대결 경기 결과 목록 보기"):
            h2h_disp = h2h.copy()
            h2h_disp['Score'] = h2h_disp['FTHG'].astype(int).astype(str) + " : " + h2h_disp['FTAG'].astype(int).astype(str)
            h2h_disp['Date'] = h2h_disp['Date'].dt.strftime('%Y-%m-%d')
            st.dataframe(h2h_disp[['Date', 'Season', 'HomeTeam', 'Score', 'AwayTeam', 'Referee']].rename(
                columns={'Date':'날짜', 'Season':'시즌', 'HomeTeam':'홈팀', 'Score':'스코어', 'AwayTeam':'원정팀', 'Referee':'주심'}
            ), use_container_width=True, hide_index=True)
    else:
        st.info("두 팀 간의 맞대결 기록이 없습니다.")
        
    st.markdown("---")
    
    st.markdown("##### 📈 **양 팀 최근 5경기 흐름 (Form)**")
    col_f1, col_f2 = st.columns(2)
    
    with col_f1:
        st.markdown(f"🏠 **{home_team} 최근 5경기 (홈/원정 포함)**")
        h_recent = df[(df['HomeTeam'] == home_team) | (df['AwayTeam'] == home_team)].head(5).copy()
        
        if not h_recent.empty:
            def get_res(row, team):
                if row['HomeTeam'] == team:
                    return '승' if row['FTR'] == 'H' else ('무' if row['FTR'] == 'D' else '패')
                else:
                    return '승' if row['FTR'] == 'A' else ('무' if row['FTR'] == 'D' else '패')
            
            h_recent['Result'] = h_recent.apply(lambda r: get_res(r, home_team), axis=1)
            h_recent['Score'] = h_recent['FTHG'].astype(int).astype(str) + " : " + h_recent['FTAG'].astype(int).astype(str)
            h_recent['DateStr'] = h_recent['Date'].dt.strftime('%Y-%m-%d')
            
            res_str = " ".join([f"[{r}]" for r in h_recent['Result'].tolist()])
            st.markdown(f"**전적:** {res_str}")
            st.dataframe(h_recent[['DateStr', 'HomeTeam', 'Score', 'AwayTeam', 'Result']].rename(
                columns={'DateStr':'날짜', 'HomeTeam':'홈팀', 'Score':'스코어', 'AwayTeam':'원정팀', 'Result':'결과'}
            ), use_container_width=True, hide_index=True)
            
    with col_f2:
        st.markdown(f"🚀 **{away_team} 최근 5경기 (홈/원정 포함)**")
        a_recent = df[(df['HomeTeam'] == away_team) | (df['AwayTeam'] == away_team)].head(5).copy()
        
        if not a_recent.empty:
            a_recent['Result'] = a_recent.apply(lambda r: get_res(r, away_team), axis=1)
            a_recent['Score'] = a_recent['FTHG'].astype(int).astype(str) + " : " + a_recent['FTAG'].astype(int).astype(str)
            a_recent['DateStr'] = a_recent['Date'].dt.strftime('%Y-%m-%d')
            
            res_str_a = " ".join([f"[{r}]" for r in a_recent['Result'].tolist()])
            st.markdown(f"**전적:** {res_str_a}")
            st.dataframe(a_recent[['DateStr', 'HomeTeam', 'Score', 'AwayTeam', 'Result']].rename(
                columns={'DateStr':'날짜', 'HomeTeam':'홈팀', 'Score':'스코어', 'AwayTeam':'원정팀', 'Result':'결과'}
            ), use_container_width=True, hide_index=True)

# ------------------------------------------
# Page 2: 배당 직접 입력 & 수율 분석
# ------------------------------------------
with tab2:
    st.subheader(f"💰 현재 경기 배당 직접 입력 및 과거 유사 배당 적중 분석")
    st.caption("이번 경기의 실제 배당을 아래에 입력해 보세요. 과거 리그 데이터에서 해당 배당을 받았을 때의 실제 적중률을 즉시 계산합니다.")
    
    # 1. 배당 직접 입력 창
    col_in1, col_in2, col_in3, col_in4 = st.columns(4)
    with col_in1:
        input_h_odds = st.number_input(f"🏠 {home_team} 홈 승리 배당", min_value=1.01, max_value=20.0, value=1.75, step=0.05)
    with col_in2:
        input_d_odds = st.number_input("🤝 무승부 배당", min_value=1.01, max_value=20.0, value=3.40, step=0.05)
    with col_in3:
        input_a_odds = st.number_input(f"🚀 {away_team} 원정 승리 배당", min_value=1.01, max_value=20.0, value=4.50, step=0.05)
    with col_in4:
        tolerance = st.slider("배당 허용 오차 범위 (±)", min_value=0.02, max_value=0.30, value=0.10, step=0.02)
    
    st.markdown("---")
    
    # 2. 입력 배당 기반 실시간 필터링
    if 'B365H' in df.columns:
        # A. 리그 전체에서 입력한 홈 배당과 유사한 경기들
        similar_league_games = df[
            (df['B365H'] >= input_h_odds - tolerance) & 
            (df['B365H'] <= input_h_odds + tolerance)
        ].copy()
        
        # B. 선택한 홈팀의 홈 경기 중 입력한 배당과 유사한 경기들
        similar_home_team_games = df[
            (df['HomeTeam'] == home_team) & 
            (df['B365H'] >= input_h_odds - tolerance) & 
            (df['B365H'] <= input_h_odds + tolerance)
        ].copy()
        
        st.markdown(f"#### 📊 입력 배당 [ **{input_h_odds:.2f}** (±{tolerance:.2f}) ] 과거 적중 분석 결과")
        
        # 리그 전체 유사 배당 성적
        if not similar_league_games.empty:
            tot_cnt = len(similar_league_games)
            h_win = len(similar_league_games[similar_league_games['FTR'] == 'H'])
            draw = len(similar_league_games[similar_league_games['FTR'] == 'D'])
            a_win = len(similar_league_games[similar_league_games['FTR'] == 'A'])
            
            similar_league_games['TotalGoals'] = similar_league_games['FTHG'] + similar_league_games['FTAG']
            o25 = len(similar_league_games[similar_league_games['TotalGoals'] > 2.5])
            
            st.markdown(f"##### 1️⃣ **{selected_league} 전체** 유사 배당 성적 (`총 {tot_cnt}경기`)")
            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.metric("홈 승리 적중률", f"{(h_win/tot_cnt)*100:.1f}%", f"{h_win} / {tot_cnt} 경기")
            mc2.metric("무승부 발생률", f"{(draw/tot_cnt)*100:.1f}%", f"{draw} / {tot_cnt} 경기")
            mc3.metric("원정 승리 발생률", f"{(a_win/tot_cnt)*100:.1f}%", f"{a_win} / {tot_cnt} 경기")
            mc4.metric("2.5 골 오버 비율", f"{(o25/tot_cnt)*100:.1f}%", f"{o25} / {tot_cnt} 경기")
            
            st.markdown("---")
            
            # 특정 홈팀의 유사 배당 성적
            st.markdown(f"##### 2️⃣ **{home_team}** 이 홈에서 동일 배당 받았을 때 성적 (`총 {len(similar_home_team_games)}경기`)")
            if not similar_home_team_games.empty:
                ht_tot = len(similar_home_team_games)
                ht_h_win = len(similar_home_team_games[similar_home_team_games['FTR'] == 'H'])
                ht_draw = len(similar_home_team_games[similar_home_team_games['FTR'] == 'D'])
                ht_a_win = len(similar_home_team_games[similar_home_team_games['FTR'] == 'A'])
                
                hc1, hc2, hc3 = st.columns(3)
                hc1.metric(f"{home_team} 승리 비율", f"{(ht_h_win/ht_tot)*100:.1f}%", f"{ht_h_win}회 승리")
                hc2.metric("무승부 비율", f"{(ht_draw/ht_tot)*100:.1f}%", f"{ht_draw}회 무승부")
                hc3.metric("패배 비율", f"{(ht_a_win/ht_tot)*100:.1f}%", f"{ht_a_win}회 패배")
                
                with st.expander(f"📋 {home_team}의 해당 배당 과거 경기 목록 보기"):
                    similar_home_team_games['Score'] = similar_home_team_games['FTHG'].astype(int).astype(str) + " : " + similar_home_team_games['FTAG'].astype(int).astype(str)
                    similar_home_team_games['DateStr'] = similar_home_team_games['Date'].dt.strftime('%Y-%m-%d')
                    st.dataframe(similar_home_team_games[['DateStr', 'Season', 'AwayTeam', 'B365H', 'B365D', 'B365A', 'Score', 'FTR']].rename(
                        columns={'DateStr':'날짜', 'Season':'시즌', 'AwayTeam':'상대원정팀', 'B365H':'홈배당', 'B365D':'무배당', 'B365A':'원정배당', 'Score':'스코어', 'FTR':'결과'}
                    ), use_container_width=True, hide_index=True)
            else:
                st.info(f"{home_team}이 홈에서 해당 배당범위[{input_h_odds - tolerance:.2f} ~ {input_h_odds + tolerance:.2f}]를 받아 경기를 치른 과거 기록이 없습니다.")
        else:
            st.warning("입력하신 배당 범위에 해당하는 유사 경기 데이터가 없습니다. 허용 오차 범위를 조금 늘려보세요.")
    else:
        st.info("이 리그 데이터에는 배당 수치가 포함되어 있지 않습니다.")

# ------------------------------------------
# Page 3: 매치 심판 성향
# ------------------------------------------
with tab3:
    st.subheader(f"🟨 심판 판정 성향 및 팀별 영향 분석")
    
    if 'Referee' in df.columns:
        available_refs = sorted(df['Referee'].dropna().unique().tolist())
        default_ref = h2h.iloc[0]['Referee'] if not h2h.empty and pd.notna(h2h.iloc[0]['Referee']) else available_refs[0]
        
        selected_ref = st.selectbox("조회할 주심을 선택하세요:", available_refs, index=available_refs.index(default_ref) if default_ref in available_refs else 0)
        
        ref_matches = df[df['Referee'] == selected_ref].copy()
        st.markdown(f"##### 👨‍⚖️ **{selected_ref}** 주심 통계 (`총 {len(ref_matches)}경기 관장`)")
        
        if not ref_matches.empty:
            ref_matches['TotalCards'] = ref_matches['HY'].fillna(0) + ref_matches['AY'].fillna(0) + (ref_matches['HR'].fillna(0) + ref_matches['AR'].fillna(0))*2
            ref_matches['TotalFouls'] = ref_matches['HF'].fillna(0) + ref_matches['AF'].fillna(0)
            
            avg_cards = ref_matches['TotalCards'].mean()
            avg_fouls = ref_matches['TotalFouls'].mean()
            home_card_ratio = (ref_matches['HY'].sum() / (ref_matches['HY'].sum() + ref_matches['AY'].sum() + 1e-5)) * 100
            
            rc1, rc2, rc3 = st.columns(3)
            rc1.metric("경기당 평균 카드 수", f"{avg_cards:.2f} 개")
            rc2.metric("경기당 평균 파울 수", f"{avg_fouls:.1f} 회")
            rc3.metric("홈팀 경고 부여 비율", f"{home_card_ratio:.1f}%", "50%보다 높으면 홈팀에 엄격")
            
            st.markdown("---")
            st.markdown(f"##### 🔍 **{selected_ref} 주심이 진행한 {home_team} 및 {away_team} 경기 기록**")
            
            team_ref_matches = ref_matches[(ref_matches['HomeTeam'].isin([home_team, away_team])) | (ref_matches['AwayTeam'].isin([home_team, away_team]))]
            
            if not team_ref_matches.empty:
                team_ref_matches['Score'] = team_ref_matches['FTHG'].astype(int).astype(str) + " : " + team_ref_matches['FTAG'].astype(int).astype(str)
                team_ref_matches['Cards'] = "홈 " + team_ref_matches['HY'].fillna(0).astype(int).astype(str) + " / 원정 " + team_ref_matches['AY'].fillna(0).astype(int).astype(str)
                team_ref_matches['DateStr'] = team_ref_matches['Date'].dt.strftime('%Y-%m-%d')
                
                st.dataframe(team_ref_matches[['DateStr', 'HomeTeam', 'Score', 'AwayTeam', 'Cards', 'HF', 'AF']].rename(
                    columns={'DateStr':'날짜', 'HomeTeam':'홈팀', 'Score':'스코어', 'AwayTeam':'원정팀', 'Cards':'옐로카드(홈/원정)', 'HF':'홈파울', 'AF':'원정파울'}
                ), use_container_width=True, hide_index=True)
            else:
                st.info(f"{selected_ref} 주심이 {home_team} 또는 {away_team}의 경기를 맡은 최근 기록이 없습니다.")
    else:
        st.info("심판 데이터가 제공되지 않습니다.")

# ------------------------------------------
# Page 4: 홈 vs 원정 직관적 지표 비교 (완전 개편)
# ------------------------------------------
with tab4:
    st.subheader(f"📈 {home_team} (홈 성적) vs {away_team} (원정 성적) 직관 지표 비교")
    st.caption("복잡한 통계 대신, 홈팀의 '홈 경기 성적'과 원정팀의 '원정 경기 성적'을 1:1 직관 지표로 명확하게 보여줍니다.")
    
    home_only = df[df['HomeTeam'] == home_team]
    away_only = df[df['AwayTeam'] == away_team]
    
    if not home_only.empty and not away_only.empty:
        # 주요 평균 지표 계산
        h_goals = home_only['FTHG'].mean()
        h_conceded = home_only['FTAG'].mean()
        h_shots = home_only['HS'].mean() if 'HS' in home_only.columns else 0
        h_shots_target = home_only['HST'].mean() if 'HST' in home_only.columns else 0
        h_corners = home_only['HC'].mean() if 'HC' in home_only.columns else 0
        
        a_goals = away_only['FTAG'].mean()
        a_conceded = away_only['FTHG'].mean()
        a_shots = away_only['AS'].mean() if 'AS' in away_only.columns else 0
        a_shots_target = away_only['AST'].mean() if 'AST' in away_only.columns else 0
        a_corners = away_only['AC'].mean() if 'AC' in away_only.columns else 0
        
        # 1. 득점 & 실점 비교 카드
        st.markdown("##### ⚽ **1. 경기당 평균 득점 & 실점 비교**")
        sc1, sc2, sc3, sc4 = st.columns(4)
        sc1.metric(f"🏠 {home_team} 홈 득점", f"{h_goals:.2f} 골")
        sc2.metric(f"🏠 {home_team} 홈 실점", f"{h_conceded:.2f} 골")
        sc3.metric(f"🚀 {away_team} 원정 득점", f"{a_goals:.2f} 골")
        sc4.metric(f"🚀 {away_team} 원정 실점", f"{a_conceded:.2f} 골")
        
        st.markdown("---")
        
        # 2. 슈팅 & 코너킥 직관 비교
        st.markdown("##### 🎯 **2. 경기당 공격 지표 직관 비교 (슈팅 & 코너킥)**")
        
        comp_metrics = [
            ("평균 슈팅 수", h_shots, a_shots, "개"),
            ("평균 유효 슈팅 수", h_shots_target, a_shots_target, "개"),
            ("평균 코너킥 획득 수", h_corners, a_corners, "개")
        ]
        
        for label, h_val, a_val, unit in comp_metrics:
            col_l, col_m, col_r = st.columns([1, 2, 1])
            with col_l:
                st.markdown(f"<h4 style='text-align: right; color: #38bdf8;'>🏠 {home_team}<br><b>{h_val:.1f} {unit}</b></h4>", unsafe_allow_html=True)
            with col_m:
                st.markdown(f"<p style='text-align: center; margin-bottom: 2px; color: #94a3b8;'><b>{label}</b></p>", unsafe_allow_html=True)
                # 시각적 게이지 바
                tot = h_val + a_val + 1e-5
                h_pct = (h_val / tot) * 100
                st.progress(int(h_pct))
            with col_r:
                st.markdown(f"<h4 style='text-align: left; color: #a855f7;'>🚀 {away_team}<br><b>{a_val:.1f} {unit}</b></h4>", unsafe_allow_html=True)
            st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)
            
    else:
        st.info("해당 팀들의 홈/원정 경기 데이터가 충분하지 않습니다.")