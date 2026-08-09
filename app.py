import streamlit as st
import pandas as pd
import numpy as np
import glob
import plotly.express as px
import plotly.graph_objects as go
import os
from scipy.stats import poisson

# ==========================================
# 1. 페이지 기본 설정 및 모던 다크 디자인 CSS
# ==========================================
st.set_page_config(
    page_title="맞춤형 축구 배팅 분석 대시보드",
    page_icon="⚽",
    layout="wide"
)

st.markdown("""
<style>
    /* 전체 배경 */
    .main, .stApp { background-color: #0e1726 !important; }
    
    /* 상단 지표 카드 (Metric) */
    .stMetric {
        background-color: #1b2e4b !important;
        border: 1px solid #3b3f5c !important;
        padding: 12px !important;
        border-radius: 10px !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.3);
    }
    .stMetric label { color: #888ea8 !important; font-size: 0.85rem !important; }
    .stMetric [data-testid="stMetricValue"] { color: #009688 !important; font-weight: bold !important; font-size: 1.4rem !important; }
    
    /* 탭(Tab) 스타일 */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #1b2e4b;
        border-radius: 8px;
        padding: 8px 16px;
        color: #888ea8;
        border: 1px solid #3b3f5c;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2196f3 !important;
        color: #ffffff !important;
        font-weight: bold;
    }
    
    /* 접기 상자 (Expander) 다크 테마 */
    .streamlit-expanderHeader, [data-testid="stExpander"] {
        background-color: #1b2e4b !important;
        color: #bfc9d4 !important;
        border: 1px solid #3b3f5c !important;
        border-radius: 8px !important;
    }
    
    /* 데이터프레임(표) 다크 스타일 적용 */
    [data-testid="stDataFrame"] {
        background-color: #1b2e4b !important;
        border-radius: 8px !important;
        border: 1px solid #3b3f5c !important;
        padding: 5px;
    }
    .stDataFrame table {
        color: #bfc9d4 !important;
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
    
    # 22개 전체 유럽 리그 시트 매핑
    league_sheets = {
        # 🏴󠁧󠁢󠁥󠁮󠁧󠁿 잉글랜드
        'EPL (잉글랜드 1부)': 'E0',
        '챔피언십 (잉글랜드 2부)': 'E1',
        '리그 원 (잉글랜드 3부)': 'E2',
        '리그 투 (잉글랜드 4부)': 'E3',
        '내셔널리그 (잉글랜드 5부)': 'EC',
        
        # 🏴󠁧󠁢󠁳󠁣󠁴󠁿 스코틀랜드
        '스코티시 프리미어쉽 (스코틀랜드 1부)': 'SC0',
        '스코티시 챔피언십 (스코틀랜드 2부)': 'SC1',
        '스코티시 리그 1 (스코틀랜드 3부)': 'SC2',
        '스코티시 리그 2 (스코틀랜드 4부)': 'SC3',
        
        # 🇩🇪 독일
        '분데스리가 (독일 1부)': 'D1',
        '분데스리가 2 (독일 2부)': 'D2',
        
        # 🇪🇸 스페인
        '라리가 (스페인 1부)': 'SP1',
        '세군다 디비시온 (스페인 2부)': 'SP2',
        
        # 🇮🇹 이탈리아
        '세리에 A (이탈리아 1부)': 'I1',
        '세리에 B (이탈리아 2부)': 'I2',
        
        # 🇫🇷 프랑스
        '리그 앙 (프랑스 1부)': 'F1',
        '리그 두 (프랑스 2부)': 'F2',
        
        # 🇳🇱 🇧🇪 🇵🇹 🇹🇷 🇬🇷 기타 주요 리그
        '에레디비시 (네덜란드 1부)': 'N1',
        '주필러 프로 리그 (벨기에 1부)': 'B1',
        '프리메이라리가 (포르투갈 1부)': 'P1',
        '쉬페르리그 (터키 1부)': 'T1',
        '수페르리가 엘라다 (그리스 1부)': 'G1'
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

st.title("⚽ 선택 매치업 종합 배팅 분석 대시보드")
st.caption("AI 승률/언오버/카드 예측 및 전/후반 득점, 마감 배당까지 정밀 분석이 가능한 종합 시스템입니다.")

if not league_dict:
    st.error("❌ 폴더 내 엑셀(.xlsx) 파일이 없거나 경로를 확인해 주세요.")
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
# 4. 포아송 및 카드 지표 계산 함수
# ==========================================
def calculate_poisson_probabilities(df_league, home_team, away_team):
    avg_home_goals = df_league['FTHG'].mean()
    avg_away_goals = df_league['FTAG'].mean()
    
    home_matches = df_league[df_league['HomeTeam'] == home_team]
    away_matches = df_league[df_league['AwayTeam'] == away_team]
    
    if home_matches.empty or away_matches.empty:
        return None
    
    home_attack = home_matches['FTHG'].mean() / (avg_home_goals + 1e-5)
    home_defense = home_matches['FTAG'].mean() / (avg_away_goals + 1e-5)
    
    away_attack = away_matches['FTAG'].mean() / (avg_away_goals + 1e-5)
    away_defense = away_matches['FTHG'].mean() / (avg_home_goals + 1e-5)
    
    expected_home_goals = home_attack * away_defense * avg_home_goals
    expected_away_goals = away_attack * home_defense * avg_away_goals
    
    max_goals = 6
    p_home = [poisson.pmf(i, expected_home_goals) for i in range(max_goals)]
    p_away = [poisson.pmf(i, expected_away_goals) for i in range(max_goals)]
    
    matrix = np.outer(p_home, p_away)
    
    prob_home_win = np.sum(np.tril(matrix, -1))
    prob_draw = np.sum(np.diag(matrix))
    prob_away_win = np.sum(np.triu(matrix, 1))
    
    prob_over25 = 0.0
    for i in range(max_goals):
        for j in range(max_goals):
            if i + j > 2.5:
                prob_over25 += matrix[i, j]
                
    # 카드 지표 계산 (HY: 홈옐로, AY: 원정옐로)
    h_cards = (home_matches['HY'].mean() if 'HY' in home_matches.columns else 0)
    a_cards = (away_matches['AY'].mean() if 'AY' in away_matches.columns else 0)
    exp_total_cards = h_cards + a_cards
    
    return {
        'h_attack': home_attack,
        'h_defense': home_defense,
        'a_attack': away_attack,
        'a_defense': away_defense,
        'exp_h_goals': expected_home_goals,
        'exp_a_goals': expected_away_goals,
        'home_win': prob_home_win * 100,
        'draw': prob_draw * 100,
        'away_win': prob_away_win * 100,
        'over25': prob_over25 * 100,
        'under25': (1 - prob_over25) * 100,
        'fair_home_odds': 1 / (prob_home_win + 1e-5),
        'fair_draw_odds': 1 / (prob_draw + 1e-5),
        'fair_away_odds': 1 / (prob_away_win + 1e-5),
        'h_avg_cards': h_cards,
        'a_avg_cards': a_cards,
        'exp_cards': exp_total_cards
    }

ai_result = calculate_poisson_probabilities(df, home_team, away_team)

# ==========================================
# 5. 5개 페이지 구성 (Page 1 = AI 종합 예측)
# ==========================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🤖 Page 1: AI 종합 예측 & 카드 시뮬레이션",
    "⚔️ Page 2: 맞대결 전적 & 전/후반 득점",
    "💰 Page 3: 초기 vs 마감 배당 세부 분석",
    "🟨 Page 4: 매치 심판 상세 판정 성향",
    "📈 Page 5: 홈 vs 원정 유효슈팅 & 지표"
])

# ------------------------------------------
# Page 1: AI 종합 예측 & 카드 시뮬레이션 (신규 첫 페이지)
# ------------------------------------------
with tab1:
    st.subheader(f"🤖 AI 종합 예측 리포트 ({home_team} vs {away_team})")
    st.caption("포아송 분포 알고리즘과 리그 통계를 기반으로 승무패, 언오버, 기대득점 및 카드 발생 위험도를 시뮬레이션합니다.")
    
    if ai_result:
        # 1. AI 승률 계산 및 적정 배당
        st.markdown("##### 🎯 **1. AI 승무패 예상 확률 및 적정 배당 (True Odds)**")
        c_ai1, c_ai2, c_ai3, c_ai4, c_ai5 = st.columns(5)
        c_ai1.metric(f"🏠 {home_team} 승리 확률", f"{ai_result['home_win']:.1f}%", f"적정 {ai_result['fair_home_odds']:.2f}")
        c_ai2.metric("🤝 무승부 확률", f"{ai_result['draw']:.1f}%", f"적정 {ai_result['fair_draw_odds']:.2f}")
        c_ai3.metric(f"🚀 {away_team} 승리 확률", f"{ai_result['away_win']:.1f}%", f"적정 {ai_result['fair_away_odds']:.2f}")
        c_ai4.metric("🔥 2.5 오버 확률", f"{ai_result['over25']:.1f}%", f"적정 {100/(ai_result['over25']+1e-5):.2f}")
        c_ai5.metric("🛡️ 2.5 언더 확률", f"{ai_result['under25']:.1f}%", f"적정 {100/(ai_result['under25']+1e-5):.2f}")
        
        st.markdown("---")
        
        # 2. 전력 지수 및 기대 득점(xG)
        st.markdown("##### 📊 **2. 팀별 홈/원정 전력 지수 & AI 예상 기대득점 (xG)**")
        p1, p2, p3, p4, p5 = st.columns(5)
        p1.metric(f"🏠 {home_team} 홈 공격력", f"{ai_result['h_attack']:.2f}", "1.0 이상 = 우수")
        p2.metric(f"🏠 {home_team} 홈 수비력", f"{ai_result['h_defense']:.2f}", "1.0 이하 = 우수")
        p3.metric(f"🚀 {away_team} 원정 공격력", f"{ai_result['a_attack']:.2f}", "1.0 이상 = 우수")
        p4.metric(f"🚀 {away_team} 원정 수비력", f"{ai_result['a_defense']:.2f}", "1.0 이하 = 우수")
        p5.metric("⚽ AI 예상 스코어 (xG)", f"{ai_result['exp_h_goals']:.2f} : {ai_result['exp_a_goals']:.2f}", f"합계 {ai_result['exp_h_goals']+ai_result['exp_a_goals']:.2f}골")
        
        st.markdown("---")
        
        # 3. AI 카드(경고/퇴장) 시뮬레이션 (추가)
        st.markdown("##### 🟨 **3. AI 매치 카드(Yellow Cards) 위험도 시뮬레이션**")
        kc1, kc2, kc3 = st.columns(3)
        kc1.metric(f"🏠 {home_team} 홈 경기당 카드 수집", f"{ai_result['h_avg_cards']:.2f} 개")
        kc2.metric(f"🚀 {away_team} 원정 경기당 카드 수집", f"{ai_result['a_avg_cards']:.2f} 개")
        
        card_risk = "높음 (격렬한 경기 예상)" if ai_result['exp_cards'] >= 4.5 else "보통 (평이한 수준)"
        kc3.metric("🚨 예상 총 카드 수 (Card Market)", f"약 {ai_result['exp_cards']:.1f} 개", f"위험도: {card_risk}")
    else:
        st.info("AI 예측을 위한 양 팀의 최소 경기 데이터가 부족합니다.")

# ------------------------------------------
# Page 2: 맞대결 전적 & 전/후반 득점
# ------------------------------------------
with tab2:
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
        
        st.markdown("---")
        if 'HTHG' in h2h.columns and 'HTAG' in h2h.columns:
            st.markdown("##### ⏱️ **맞대결 전반전 vs 후반전 득점 분포**")
            h2h['HT_Goals'] = h2h['HTHG'] + h2h['HTAG']
            h2h['2H_Goals'] = h2h['TotalGoals'] - h2h['HT_Goals']
            
            avg_ht = h2h['HT_Goals'].mean()
            avg_2h = h2h['2H_Goals'].mean()
            ht_over05 = (len(h2h[h2h['HT_Goals'] > 0.5]) / len(h2h)) * 100
            
            fc1, fc2, fc3 = st.columns(3)
            fc1.metric("전반전 평균 득점", f"{avg_ht:.2f} 골")
            fc2.metric("후반전 평균 득점", f"{avg_2h:.2f} 골")
            fc3.metric("전반전 0.5 오버 확률", f"{ht_over05:.1f}%")
        
        with st.expander("📄 최근 맞대결 경기 스코어 (전반 스코어 포함) 목록 보기"):
            h2h_disp = h2h.copy()
            h2h_disp['Score'] = h2h_disp['FTHG'].astype(int).astype(str) + " : " + h2h_disp['FTAG'].astype(int).astype(str)
            if 'HTHG' in h2h_disp.columns:
                h2h_disp['HT_Score'] = h2h_disp['HTHG'].fillna(0).astype(int).astype(str) + " : " + h2h_disp['HTAG'].fillna(0).astype(int).astype(str)
            else:
                h2h_disp['HT_Score'] = "-"
            
            h2h_disp['Date'] = h2h_disp['Date'].dt.strftime('%Y-%m-%d')
            
            disp_cols = ['Date', 'Season', 'HomeTeam', 'HT_Score', 'Score', 'AwayTeam']
            renames = {'Date':'날짜', 'Season':'시즌', 'HomeTeam':'홈팀', 'HT_Score':'전반스코어', 'Score':'최종스코어', 'AwayTeam':'원정팀'}
            if 'Referee' in h2h_disp.columns:
                disp_cols.append('Referee')
                renames['Referee'] = '주심'
                
            st.dataframe(h2h_disp[disp_cols].rename(columns=renames), use_container_width=True, hide_index=True)
    else:
        st.info("두 팀 간의 맞대결 기록이 없습니다.")
        
    st.markdown("---")
    
    st.markdown("##### 📈 **양 팀 최근 5경기 흐름 (Form)**")
    col_f1, col_f2 = st.columns(2)
    
    def get_res(row, team):
        if row['HomeTeam'] == team:
            return '승' if row['FTR'] == 'H' else ('무' if row['FTR'] == 'D' else '패')
        else:
            return '승' if row['FTR'] == 'A' else ('무' if row['FTR'] == 'D' else '패')
            
    with col_f1:
        st.markdown(f"🏠 **{home_team} 최근 5경기 (홈/원정 포함)**")
        h_recent = df[(df['HomeTeam'] == home_team) | (df['AwayTeam'] == home_team)].head(5).copy()
        
        if not h_recent.empty:
            h_recent['Result'] = h_recent.apply(lambda r: get_res(r, home_team), axis=1)
            h_recent['Score'] = h_recent['FTHG'].astype(int).astype(str) + " : " + h_recent['FTAG'].astype(int).astype(str)
            h_recent['DateStr'] = h_recent['Date'].dt.strftime('%Y-%m-%d')
            
            res_str = " ".join([f"[{r}]" for r in h_recent['Result'].tolist()])
            st.markdown(f"**최근 전적:** `{res_str}`")
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
            st.markdown(f"**최근 전적:** `{res_str_a}`")
            st.dataframe(a_recent[['DateStr', 'HomeTeam', 'Score', 'AwayTeam', 'Result']].rename(
                columns={'DateStr':'날짜', 'HomeTeam':'홈팀', 'Score':'스코어', 'AwayTeam':'원정팀', 'Result':'결과'}
            ), use_container_width=True, hide_index=True)

# ------------------------------------------
# Page 3: 초기 vs 마감 배당 세부 분석
# ------------------------------------------
with tab3:
    st.subheader(f"💰 초기 배당 vs 마감 배당(Closing Odds) 세부 분석")
    st.caption("이번 경기의 실제 배당을 입력하여 과거 유사 배당의 적중률을 분석하고, 초기 배당과 마감 배당(홈/무/원정) 변동을 한눈에 비교합니다.")
    
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
    
    if 'B365H' in df.columns:
        similar_league_games = df[
            (df['B365H'] >= input_h_odds - tolerance) & 
            (df['B365H'] <= input_h_odds + tolerance)
        ].copy()
        
        st.markdown(f"#### 📊 입력 배당 [ 홈 **{input_h_odds:.2f}** / 무 **{input_d_odds:.2f}** / 원정 **{input_a_odds:.2f}** (오차 ±{tolerance:.2f}) ] 적중 성적")
        
        if not similar_league_games.empty:
            tot_cnt = len(similar_league_games)
            h_win = len(similar_league_games[similar_league_games['FTR'] == 'H'])
            draw = len(similar_league_games[similar_league_games['FTR'] == 'D'])
            a_win = len(similar_league_games[similar_league_games['FTR'] == 'A'])
            
            has_close = ('B365CH' in similar_league_games.columns) and ('B365CD' in similar_league_games.columns) and ('B365CA' in similar_league_games.columns)
            if has_close:
                drop_h = similar_league_games[similar_league_games['B365CH'] < similar_league_games['B365H']]
                drop_h_win = len(drop_h[drop_h['FTR'] == 'H']) if len(drop_h) > 0 else 0
                drop_win_rate = (drop_h_win / len(drop_h) * 100) if len(drop_h) > 0 else 0
            
            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.metric("홈 승리 적중률", f"{(h_win/tot_cnt)*100:.1f}%", f"{h_win} / {tot_cnt} 경기")
            mc2.metric("무승부 발생률", f"{(draw/tot_cnt)*100:.1f}%", f"{draw} / {tot_cnt} 경기")
            mc3.metric("원정 승리 발생률", f"{(a_win/tot_cnt)*100:.1f}%", f"{a_win} / {tot_cnt} 경기")
            
            if has_close:
                mc4.metric("마감 시 홈배당 하락 시 승률", f"{drop_win_rate:.1f}%", f"{len(drop_h)}경기 중 하락")
            else:
                similar_league_games['TotalGoals'] = similar_league_games['FTHG'] + similar_league_games['FTAG']
                o25 = len(similar_league_games[similar_league_games['TotalGoals'] > 2.5])
                mc4.metric("2.5 골 오버 비율", f"{(o25/tot_cnt)*100:.1f}%")
                
            st.markdown("---")
            
            st.markdown("##### 📋 **과거 유사 배당 경기 (초기 배당 vs 마감 배당 전체 목록)**")
            similar_league_games['Score'] = similar_league_games['FTHG'].astype(int).astype(str) + " : " + similar_league_games['FTAG'].astype(int).astype(str)
            similar_league_games['DateStr'] = similar_league_games['Date'].dt.strftime('%Y-%m-%d')
            
            show_cols = ['DateStr', 'HomeTeam', 'AwayTeam', 'B365H', 'B365D', 'B365A']
            renames = {
                'DateStr': '날짜', 'HomeTeam': '홈팀', 'AwayTeam': '원정팀',
                'B365H': '초기 [홈]', 'B365D': '초기 [무]', 'B365A': '초기 [원정]'
            }
            
            if has_close:
                show_cols.extend(['B365CH', 'B365CD', 'B365CA'])
                renames['B365CH'] = '마감 [홈]'
                renames['B365CD'] = '마감 [무]'
                renames['B365CA'] = '마감 [원정]'
                
            show_cols.extend(['Score', 'FTR'])
            renames['Score'] = '스코어'
            renames['FTR'] = '결과'
            
            st.dataframe(
                similar_league_games[show_cols].rename(columns=renames),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.warning("입력하신 배당 범위에 해당하는 유사 경기 데이터가 없습니다.")
    else:
        st.info("이 리그 데이터에는 배당 수치가 포함되어 있지 않습니다.")

# ------------------------------------------
# Page 4: 매치 심판 성향
# ------------------------------------------
with tab4:
    st.subheader(f"🟨 심판 판정 성향 및 팀별 영향 분석")
    
    if 'Referee' in df.columns:
        available_refs = sorted(df['Referee'].dropna().unique().tolist())
        default_ref = h2h.iloc[0]['Referee'] if not h2h.empty and 'Referee' in h2h.columns and pd.notna(h2h.iloc[0]['Referee']) else (available_refs[0] if available_refs else None)
        
        if available_refs:
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
            st.info("선택 가능한 주심 목록이 없습니다.")
    else:
        st.info("💡 라리가, 세리에A, 분데스리가 데이터에는 주심(Referee) 정보가 포함되어 있지 않습니다. (EPL 전용 기능)")

# ------------------------------------------
# Page 5: 홈 vs 원정 유효슈팅 & 지표 분석
# ------------------------------------------
with tab5:
    st.subheader(f"📈 {home_team} (홈 성적) vs {away_team} (원정 성적) 유효슈팅 & 지표 비교")
    st.caption("홈팀의 홈 세부 지표와 원정팀의 원정 세부 지표(득점, 슈팅, 유효슈팅, 코너킥)를 1:1 대조합니다.")
    
    home_only = df[df['HomeTeam'] == home_team]
    away_only = df[df['AwayTeam'] == away_team]
    
    if not home_only.empty and not away_only.empty:
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
        
        st.markdown("##### ⚽ **1. 경기당 평균 득점 & 실점 비교**")
        sc1, sc2, sc3, sc4 = st.columns(4)
        sc1.metric(f"🏠 {home_team} 홈 득점", f"{h_goals:.2f} 골")
        sc2.metric(f"🏠 {home_team} 홈 실점", f"{h_conceded:.2f} 골")
        sc3.metric(f"🚀 {away_team} 원정 득점", f"{a_goals:.2f} 골")
        sc4.metric(f"🚀 {away_team} 원정 실점", f"{a_conceded:.2f} 골")
        
        st.markdown("---")
        
        st.markdown("##### 🎯 **2. 유효슈팅 & 슈팅 & 코너킥 1:1 직관 비교**")
        
        comp_metrics = [
            ("평균 유효 슈팅 수 (On Target)", h_shots_target, a_shots_target, "개"),
            ("평균 전체 슈팅 수 (Total Shots)", h_shots, a_shots, "개"),
            ("평균 코너킥 획득 수 (Corners)", h_corners, a_corners, "개")
        ]
        
        for label, h_val, a_val, unit in comp_metrics:
            col_l, col_m, col_r = st.columns([1, 2, 1])
            with col_l:
                st.markdown(f"<h4 style='text-align: right; color: #009688;'>🏠 {home_team}<br><b>{h_val:.1f} {unit}</b></h4>", unsafe_allow_html=True)
            with col_m:
                st.markdown(f"<p style='text-align: center; margin-bottom: 2px; color: #888ea8;'><b>{label}</b></p>", unsafe_allow_html=True)
                tot = h_val + a_val + 1e-5
                h_pct = (h_val / tot) * 100
                st.progress(int(h_pct))
            with col_r:
                st.markdown(f"<h4 style='text-align: left; color: #2196f3;'>🚀 {away_team}<br><b>{a_val:.1f} {unit}</b></h4>", unsafe_allow_html=True)
            st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)
            
    else:
        st.info("해당 팀들의 홈/원정 경기 데이터가 충분하지 않습니다.")
