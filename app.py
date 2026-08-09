import streamlit as st
import pandas as pd
import numpy as np
import glob
import plotly.express as px
import plotly.graph_objects as go
import os
import asyncio
from scipy.stats import poisson

# 비동기 understat 패키지 불러오기
try:
    import aiohttp
    from understat import Understat
    HAS_UNDERSTAT = True
except ImportError:
    HAS_UNDERSTAT = False

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
    
    /* 2줄 레이아웃 탭 스타일 */
    .stTabs [data-baseweb="tab-list"] { 
        gap: 8px; 
        flex-wrap: wrap !important;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #1b2e4b;
        border-radius: 8px;
        padding: 10px 18px;
        color: #bfc9d4;
        border: 1px solid #3b3f5c;
        margin-bottom: 5px;
        font-weight: 500;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2196f3 !important;
        color: #ffffff !important;
        font-weight: bold;
        border-color: #2196f3 !important;
    }
    
    /* 접기 상자 (Expander) 다크 테마 */
    .streamlit-expanderHeader, [data-testid="stExpander"] {
        background-color: #1b2e4b !important;
        color: #bfc9d4 !important;
        border: 1px solid #3b3f5c !important;
        border-radius: 8px !important;
    }
    
    /* 데이터프레임 다크 스타일 적용 */
    [data-testid="stDataFrame"] {
        background-color: #1b2e4b !important;
        border-radius: 8px !important;
        border: 1px solid #3b3f5c !important;
        padding: 5px;
    }
    .stDataFrame table { color: #bfc9d4 !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 데이터 로드 및 전처리 (22개 전 리그 매핑)
# ==========================================
@st.cache_data
def load_data():
    files = sorted(glob.glob("*.xlsx"))
    if not files:
        return {}
    
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

# ------------------------------------------
# 💡 비동기 understat 패키지 적용 xG 수집 함수
# ------------------------------------------
async def _async_get_understat(league_code, season_year):
    async with aiohttp.ClientSession() as session:
        understat = Understat(session)
        teams_data = await understat.get_teams(league_code, season_year)
        parsed = []
        for t in teams_data:
            history = t.get('history', [])
            total_xg = sum(float(m['xG']) for m in history) if history else 0
            total_xga = sum(float(m['xGA']) for m in history) if history else 0
            total_xpts = sum(float(m['xPTS']) for m in history) if history else 0
            gp = len(history) if history else 1
            
            parsed.append({
                'Team': t['title'],
                'GP': gp,
                'real_xG': total_xg,
                'real_xGA': total_xga,
                'avg_xG': total_xg / (gp + 1e-5),
                'avg_xGA': total_xga / (gp + 1e-5),
                'xPTS': total_xpts
            })
        return pd.DataFrame(parsed)

@st.cache_data(ttl=3600)
def fetch_understat_xg(selected_league_name):
    if not HAS_UNDERSTAT:
        return None
        
    understat_map = {
        'EPL (잉글랜드 1부)': 'epl',
        '라리가 (스페인 1부)': 'la_liga',
        '분데스리가 (독일 1부)': 'bundesliga',
        '세리에 A (이탈리아 1부)': 'serie_a',
        '리그 앙 (프랑스 1부)': 'ligue_1'
    }
    
    if selected_league_name not in understat_map:
        return None
        
    league_code = understat_map[selected_league_name]
    
    for year in [2025, 2024]:
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            df_res = loop.run_until_complete(_async_get_understat(league_code, year))
            loop.close()
            if df_res is not None and not df_res.empty:
                return df_res
        except Exception:
            pass
            
    return None

league_dict = load_data()

st.title("⚽ 선택 매치업 종합 배팅 분석 대시보드")
st.caption("AI 승률/언오버/카드/유효슈팅 종합 예측 및 최적 배팅 조건 자동 산출 시스템입니다.")

if not league_dict:
    st.error("❌ 폴더 내 엑셀(.xlsx) 파일이 없거나 경로를 확인해 주세요.")
    st.stop()

# ==========================================
# 3. 사이드바 - 리그, 매치업 및 결장자 선택
# ==========================================
st.sidebar.header("⚙️ 매치업 선택")
selected_league = st.sidebar.selectbox("1. 리그 선택", list(league_dict.keys()))

df = league_dict[selected_league]
teams = sorted(list(set(df['HomeTeam'].dropna().unique()).union(set(df['AwayTeam'].dropna().unique()))))

col_s1, col_s2 = st.sidebar.columns(2)
home_team = st.sidebar.selectbox("2. 홈 팀 선택", teams, index=0)
away_team = st.sidebar.selectbox("3. 원정 팀 선택", [t for t in teams if t != home_team], index=min(1, len(teams)-1))

st.sidebar.markdown("---")
st.sidebar.subheader("🏥 주요 결장자(부상/징계) 설정")
home_inj_att = st.sidebar.slider(f"🏠 {home_team} 주요 공격진 결장", 0, 3, 0, help="결장자 1명당 공격력 8% 감소")
home_inj_def = st.sidebar.slider(f"🏠 {home_team} 주요 수비진/키퍼 결장", 0, 3, 0, help="결장자 1명당 수비 불안정성 8% 증가")

st.sidebar.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)
away_inj_att = st.sidebar.slider(f"🚀 {away_team} 주요 공격진 결장", 0, 3, 0, help="결장자 1명당 공격력 8% 감소")
away_inj_def = st.sidebar.slider(f"🚀 {away_team} 주요 수비진/키퍼 결장", 0, 3, 0, help="결장자 1명당 수비 불안정성 8% 증가")

st.sidebar.markdown("---")
st.sidebar.success(f"🎯 **{home_team} (홈) vs {away_team} (원정)**")

# 5대 리그 실제 xG 데이터 호출
df_real_xg = fetch_understat_xg(selected_league)

# ==========================================
# 4. 포아송 및 정밀 배팅 지표 산출 알고리즘
# ==========================================
def calculate_poisson_probabilities(df_league, home_team, away_team, h_inj_att=0, h_inj_def=0, a_inj_att=0, a_inj_def=0, xg_df=None):
    avg_home_goals = df_league['FTHG'].mean()
    avg_away_goals = df_league['FTAG'].mean()
    
    home_matches = df_league[df_league['HomeTeam'] == home_team]
    away_matches = df_league[df_league['AwayTeam'] == away_team]
    
    if home_matches.empty or away_matches.empty:
        return None
    
    # 1. 기본 공격/수비 지수
    home_attack = home_matches['FTHG'].mean() / (avg_home_goals + 1e-5)
    home_defense = home_matches['FTAG'].mean() / (avg_away_goals + 1e-5)
    
    away_attack = away_matches['FTAG'].mean() / (avg_away_goals + 1e-5)
    away_defense = away_matches['FTHG'].mean() / (avg_home_goals + 1e-5)
    
    # 2. xG 보정
    if xg_df is not None and not xg_df.empty:
        h_row = xg_df[xg_df['Team'].str.contains(home_team[:4], case=False, na=False)]
        a_row = xg_df[xg_df['Team'].str.contains(away_team[:4], case=False, na=False)]
        
        if not h_row.empty and not a_row.empty:
            avg_league_xg = xg_df['avg_xG'].mean() + 1e-5
            home_attack = (h_row['avg_xG'].values[0] / avg_league_xg)
            away_attack = (a_row['avg_xG'].values[0] / avg_league_xg)
            home_defense = (h_row['avg_xGA'].values[0] / avg_league_xg)
            away_defense = (a_row['avg_xGA'].values[0] / avg_league_xg)
            
    elif 'HST' in home_matches.columns and 'AST' in away_matches.columns and not home_matches['HST'].isna().all():
        h_xg_stat = (home_matches['HST'].mean() * 0.32) + ((home_matches['HS'].mean() - home_matches['HST'].mean()) * 0.06)
        a_xg_stat = (away_matches['AST'].mean() * 0.32) + ((away_matches['AS'].mean() - away_matches['AST'].mean()) * 0.06)
        
        league_h_xg = (df_league['HST'].mean() * 0.32) + ((df_league['HS'].mean() - df_league['HST'].mean()) * 0.06)
        league_a_xg = (df_league['AST'].mean() * 0.32) + ((df_league['AS'].mean() - df_league['AST'].mean()) * 0.06)
        
        home_attack = (home_attack * 0.5) + ((h_xg_stat / (league_h_xg + 1e-5)) * 0.5)
        away_attack = (away_attack * 0.5) + ((a_xg_stat / (league_a_xg + 1e-5)) * 0.5)

    # 3. 결장자 반영
    home_attack = max(0.1, home_attack * (1.0 - h_inj_att * 0.08))
    home_defense = home_defense * (1.0 + h_inj_def * 0.08)
    away_attack = max(0.1, away_attack * (1.0 - a_inj_att * 0.08))
    away_defense = away_defense * (1.0 + a_inj_def * 0.08)
    
    # 기대 득점(xG) 계산
    expected_home_goals = home_attack * away_defense * avg_home_goals
    expected_away_goals = away_attack * home_defense * avg_away_goals
    
    max_goals = 6
    p_home = [poisson.pmf(i, expected_home_goals) for i in range(max_goals)]
    p_away = [poisson.pmf(i, expected_away_goals) for i in range(max_goals)]
    
    matrix = np.outer(p_home, p_away)
    
    prob_home_win = np.sum(np.tril(matrix, -1))
    prob_draw = np.sum(np.diag(matrix))
    prob_away_win = np.sum(np.triu(matrix, 1))
    
    prob_over25 = sum(matrix[i, j] for i in range(max_goals) for j in range(max_goals) if i + j > 2.5)
    
    # 카드 및 슈팅/코너킥 평균 지표 계산
    h_cards = (home_matches['HY'].mean() if 'HY' in home_matches.columns else 0)
    a_cards = (away_matches['AY'].mean() if 'AY' in away_matches.columns else 0)
    
    h_sot = (home_matches['HST'].mean() if 'HST' in home_matches.columns else 0)
    a_sot = (away_matches['AST'].mean() if 'AST' in away_matches.columns else 0)
    
    h_cor = (home_matches['HC'].mean() if 'HC' in home_matches.columns else 0)
    a_cor = (away_matches['AC'].mean() if 'AC' in away_matches.columns else 0)
    
    return {
        'h_attack': home_attack, 'h_defense': home_defense,
        'a_attack': away_attack, 'a_defense': away_defense,
        'exp_h_goals': expected_home_goals, 'exp_a_goals': expected_away_goals,
        'home_win': prob_home_win * 100, 'draw': prob_draw * 100, 'away_win': prob_away_win * 100,
        'over25': prob_over25 * 100, 'under25': (1 - prob_over25) * 100,
        'fair_home_odds': 1 / (prob_home_win + 1e-5),
        'fair_draw_odds': 1 / (prob_draw + 1e-5),
        'fair_away_odds': 1 / (prob_away_win + 1e-5),
        'exp_cards': h_cards + a_cards,
        'exp_sot': h_sot + a_sot,
        'exp_corners': h_cor + a_cor
    }

ai_result = calculate_poisson_probabilities(df, home_team, away_team, home_inj_att, home_inj_def, away_inj_att, away_inj_def, df_real_xg)

# ==========================================
# 5. 2줄 편의 탭(Tab) 구조 세팅
# ==========================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🤖 Page 1: AI 종합 예측 & 최적 추천",
    "⚔️ Page 2: 맞대결 전적 & 전/후반 득점",
    "💰 Page 3: 초기 vs 마감 배당 분석",
    "🟨 Page 4: 매치 심판 판정 성향",
    "📈 Page 5: 유효슈팅 & 슈팅 지표",
    "📊 Page 6: 리그 실시간 순위표",
    "🎯 Page 7: 팀별 정밀 xG 분석"
])

# ------------------------------------------
# Page 1: AI 종합 예측 & 최적 추천 조건 산출
# ------------------------------------------
with tab1:
    st.subheader(f"🤖 AI 종합 예측 & 최적 추천 조건 리포트 ({home_team} vs {away_team})")
    
    if ai_result:
        # 최적 배팅 조건 자동 예측 로직
        best_picks = []
        
        # 1) 승무패 판단
        if ai_result['home_win'] >= 58.0:
            best_picks.append(("🔥 [승무패 최고 추천]", f"🏠 {home_team} 홈 승리", f"확률 {ai_result['home_win']:.1f}% (적정배당 {ai_result['fair_home_odds']:.2f})"))
        elif ai_result['away_win'] >= 52.0:
            best_picks.append(("🔥 [승무패 최고 추천]", f"🚀 {away_team} 원정 승리", f"확률 {ai_result['away_win']:.1f}% (적정배당 {ai_result['fair_away_odds']:.2f})"))
        elif ai_result['draw'] >= 32.0:
            best_picks.append(("🤝 [승무패 추천]", "무승부 고배당 접근", f"확률 {ai_result['draw']:.1f}% (적정배당 {ai_result['fair_draw_odds']:.2f})"))
            
        # 2) 언오버 판단
        if ai_result['over25'] >= 58.0:
            best_picks.append(("⚽ [언오버 추천]", "2.5 골 오버 (Over)", f"확률 {ai_result['over25']:.1f}% (예상 합계 {ai_result['exp_h_goals']+ai_result['exp_a_goals']:.2f}골)"))
        elif ai_result['under25'] >= 58.0:
            best_picks.append(("🛡️ [언오버 추천]", "2.5 골 언더 (Under)", f"확률 {ai_result['under25']:.1f}% (예상 합계 {ai_result['exp_h_goals']+ai_result['exp_a_goals']:.2f}골)"))
            
        # 3) 카드 마켓 판단
        if ai_result['exp_cards'] >= 4.5:
            best_picks.append(("🟨 [카드 마켓 추천]", "오버 4.5 경고 카드", f"예상 카드 수 {ai_result['exp_cards']:.1f}개 (격렬한 매치 예상)"))
        elif ai_result['exp_cards'] <= 2.8:
            best_picks.append(("🟨 [카드 마켓 추천]", "언더 3.5 경고 카드", f"예상 카드 수 {ai_result['exp_cards']:.1f}개 (신사적인 매치 예상)"))
            
        # 4) 유효슈팅 및 코너킥 판단
        if ai_result['exp_sot'] >= 9.5:
            best_picks.append(("🎯 [유효슈팅 추천]", "유효슈팅 9.5 오버", f"양 팀 예상 합계 유효슈팅 {ai_result['exp_sot']:.1f}개"))
        if ai_result['exp_corners'] >= 10.0:
            best_picks.append(("🚩 [코너킥 추천]", "코너킥 9.5 오버", f"양 팀 예상 합계 코너킥 {ai_result['exp_corners']:.1f}개"))

        # 최적 조건 카드 출력
        st.markdown("##### 🏆 **AI가 분석한 이번 경기 가장 유리한 배팅 조건 (Best Value Picks)**")
        if best_picks:
            r_cols = st.columns(len(best_picks))
            for idx, (category, pick_title, detail) in enumerate(best_picks):
                with r_cols[idx]:
                    st.markdown(f"""
                    <div style='background-color:#1b2e4b; border:1px solid #009688; border-radius:10px; padding:12px;'>
                        <p style='color:#009688; margin:0; font-size:0.85rem;'><b>{category}</b></p>
                        <h4 style='color:#ffffff; margin:5px 0;'><b>{pick_title}</b></h4>
                        <p style='color:#888ea8; margin:0; font-size:0.8rem;'>{detail}</p>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("💡 양 팀 전력이 팽팽하여 승무패보다는 live 진행 상황을 확인 후 접근하는 것을 추천합니다.")
            
        st.markdown("---")
        
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
        st.markdown("##### 📊 **2. 팀별 보정 전력 지수 & AI 예상 기대득점 (xG)**")
        p1, p2, p3, p4, p5 = st.columns(5)
        p1.metric(f"🏠 {home_team} 홈 공격력", f"{ai_result['h_attack']:.2f}", "1.0 이상 = 우수")
        p2.metric(f"🏠 {home_team} 홈 수비력", f"{ai_result['h_defense']:.2f}", "1.0 이하 = 우수")
        p3.metric(f"🚀 {away_team} 원정 공격력", f"{ai_result['a_attack']:.2f}", "1.0 이상 = 우수")
        p4.metric(f"🚀 {away_team} 원정 수비력", f"{ai_result['a_defense']:.2f}", "1.0 이하 = 우수")
        p5.metric("⚽ AI 예상 스코어 (xG)", f"{ai_result['exp_h_goals']:.2f} : {ai_result['exp_a_goals']:.2f}", f"합계 {ai_result['exp_h_goals']+ai_result['exp_a_goals']:.2f}골")

# ------------------------------------------
# Page 2: 맞대결 전적 & 전/후반 득점
# ------------------------------------------
with tab2:
    st.subheader(f"⚔️ {home_team} vs {away_team} 맞대결 전적 & 최근 흐름")
    h2h = df[((df['HomeTeam'] == home_team) & (df['AwayTeam'] == away_team)) | 
             ((df['HomeTeam'] == away_team) & (df['AwayTeam'] == home_team))].sort_values(by='Date', ascending=False)
    
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
            h2h['HT_Goals'] = h2h['HTHG'] + h2h['HTAG']
            h2h['2H_Goals'] = h2h['TotalGoals'] - h2h['HT_Goals']
            
            fc1, fc2, fc3 = st.columns(3)
            fc1.metric("전반전 평균 득점", f"{h2h['HT_Goals'].mean():.2f} 골")
            fc2.metric("후반전 평균 득점", f"{h2h['2H_Goals'].mean():.2f} 골")
            fc3.metric("전반전 0.5 오버 확률", f"{(len(h2h[h2h['HT_Goals'] > 0.5])/len(h2h))*100:.1f}%")
            
        with st.expander("📄 최근 맞대결 경기 목록 보기"):
            h2h_disp = h2h.copy()
            h2h_disp['Score'] = h2h_disp['FTHG'].astype(int).astype(str) + " : " + h2h_disp['FTAG'].astype(int).astype(str)
            h2h_disp['Date'] = h2h_disp['Date'].dt.strftime('%Y-%m-%d')
            st.dataframe(h2h_disp[['Date', 'Season', 'HomeTeam', 'Score', 'AwayTeam']].rename(
                columns={'Date':'날짜', 'Season':'시즌', 'HomeTeam':'홈팀', 'Score':'최종스코어', 'AwayTeam':'원정팀'}
            ), use_container_width=True, hide_index=True)

# ------------------------------------------
# Page 3: 초기 vs 마감 배당 세부 분석
# ------------------------------------------
with tab3:
    st.subheader(f"💰 초기 배당 vs 마감 배당 세부 분석")
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
        similar_games = df[(df['B365H'] >= input_h_odds - tolerance) & (df['B365H'] <= input_h_odds + tolerance)].copy()
        if not similar_games.empty:
            tot = len(similar_games)
            hw = len(similar_games[similar_games['FTR'] == 'H'])
            dr = len(similar_games[similar_games['FTR'] == 'D'])
            aw = len(similar_games[similar_games['FTR'] == 'A'])
            
            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.metric("홈 승리 적중률", f"{(hw/tot)*100:.1f}%", f"{hw}/{tot} 경기")
            mc2.metric("무승부 발생률", f"{(dr/tot)*100:.1f}%", f"{dr}/{tot} 경기")
            mc3.metric("원정 승리 발생률", f"{(aw/tot)*100:.1f}%", f"{aw}/{tot} 경기")
            similar_games['TotalGoals'] = similar_games['FTHG'] + similar_games['FTAG']
            mc4.metric("2.5 골 오버 비율", f"{(len(similar_games[similar_games['TotalGoals']>2.5])/tot)*100:.1f}%")

# ------------------------------------------
# Page 4: 매치 심판 성향
# ------------------------------------------
with tab4:
    st.subheader(f"🟨 심판 판정 성향 분석")
    if 'Referee' in df.columns:
        available_refs = sorted(df['Referee'].dropna().unique().tolist())
        if available_refs:
            selected_ref = st.selectbox("조회할 주심을 선택하세요:", available_refs)
            ref_matches = df[df['Referee'] == selected_ref].copy()
            if not ref_matches.empty:
                ref_matches['TotalCards'] = ref_matches['HY'].fillna(0) + ref_matches['AY'].fillna(0) + (ref_matches['HR'].fillna(0) + ref_matches['AR'].fillna(0))*2
                rc1, rc2, rc3 = st.columns(3)
                rc1.metric("경기당 평균 카드 수", f"{ref_matches['TotalCards'].mean():.2f} 개")
                rc2.metric("경기당 평균 파울 수", f"{(ref_matches['HF'].fillna(0)+ref_matches['AF'].fillna(0)).mean():.1f} 회")
                rc3.metric("홈팀 경고 부여 비율", f"{(ref_matches['HY'].sum()/(ref_matches['HY'].sum()+ref_matches['AY'].sum()+1e-5))*100:.1f}%")

# ------------------------------------------
# Page 5: 유효슈팅 & 지표 분석
# ------------------------------------------
with tab5:
    st.subheader(f"📈 {home_team} (홈) vs {away_team} (원정) 슈팅 & 코너킥 세부 비교")
    home_only = df[df['HomeTeam'] == home_team]
    away_only = df[df['AwayTeam'] == away_team]
    
    if not home_only.empty and not away_only.empty:
        comp_metrics = [
            ("평균 유효 슈팅 수 (On Target)", home_only['HST'].mean() if 'HST' in home_only else 0, away_only['AST'].mean() if 'AST' in away_only else 0, "개"),
            ("평균 전체 슈팅 수 (Total Shots)", home_only['HS'].mean() if 'HS' in home_only else 0, away_only['AS'].mean() if 'AS' in away_only else 0, "개"),
            ("평균 코너킥 획득 수 (Corners)", home_only['HC'].mean() if 'HC' in home_only else 0, away_only['AC'].mean() if 'AC' in away_only else 0, "개")
        ]
        for label, h_val, a_val, unit in comp_metrics:
            col_l, col_m, col_r = st.columns([1, 2, 1])
            with col_l: st.markdown(f"<h4 style='text-align: right; color: #009688;'>🏠 {home_team}<br><b>{h_val:.1f} {unit}</b></h4>", unsafe_allow_html=True)
            with col_m:
                st.markdown(f"<p style='text-align: center; color: #888ea8;'><b>{label}</b></p>", unsafe_allow_html=True)
                st.progress(int((h_val / (h_val + a_val + 1e-5)) * 100))
            with col_r: st.markdown(f"<h4 style='text-align: left; color: #2196f3;'>🚀 {away_team}<br><b>{a_val:.1f} {unit}</b></h4>", unsafe_allow_html=True)

# ------------------------------------------
# Page 6: 리그 실시간 순위표
# ------------------------------------------
with tab6:
    st.subheader(f"📊 {selected_league} 실시간 리그 순위표")
    latest_season = df['Season'].max() if 'Season' in df.columns else None
    df_season = df[df['Season'] == latest_season] if latest_season else df
    
    standings = {t: {'GP': 0, 'W': 0, 'D': 0, 'L': 0, 'GF': 0, 'GA': 0, 'PTS': 0} for t in teams}
    for _, r in df_season.iterrows():
        ht, at, hg, ag, ftr = r['HomeTeam'], r['AwayTeam'], r['FTHG'], r['FTAG'], r['FTR']
        if ht in standings and at in standings and pd.notna(hg) and pd.notna(ag):
            standings[ht]['GP'] += 1; standings[ht]['GF'] += int(hg); standings[ht]['GA'] += int(ag)
            standings[at]['GP'] += 1; standings[at]['GF'] += int(ag); standings[at]['GA'] += int(hg)
            if ftr == 'H': standings[ht]['W'] += 1; standings[ht]['PTS'] += 3; standings[at]['L'] += 1
            elif ftr == 'A': standings[at]['W'] += 1; standings[at]['PTS'] += 3; standings[ht]['L'] += 1
            elif ftr == 'D': standings[ht]['D'] += 1; standings[ht]['PTS'] += 1; standings[at]['D'] += 1; standings[at]['PTS'] += 1

    df_rank = pd.DataFrame.from_dict(standings, orient='index').reset_index().rename(columns={'index': '팀명'})
    df_rank['GD'] = df_rank['GF'] - df_rank['GA']
    df_rank = df_rank.sort_values(by=['PTS', 'GD', 'GF'], ascending=[False, False, False]).reset_index(drop=True)
    df_rank.index = df_rank.index + 1
    df_rank['순위'] = df_rank.index
    st.dataframe(df_rank[['순위', '팀명', 'GP', 'W', 'D', 'L', 'GF', 'GA', 'GD', 'PTS']].rename(
        columns={'GP':'경기수', 'W':'승', 'D':'무', 'L':'패', 'GF':'득점', 'GA':'실점', 'GD':'득실차', 'PTS':'승점'}
    ), use_container_width=True, hide_index=True)

# ------------------------------------------
# Page 7: 팀별 정밀 xG 분석 리포트
# ------------------------------------------
with tab7:
    st.subheader(f"🎯 {selected_league} 정밀 기대득점(xG) 통계")
    if df_real_xg is not None and not df_real_xg.empty:
        df_show = df_real_xg.copy()
        df_show['xGDiff'] = df_show['real_xG'] - df_show['real_xGA']
        df_show = df_show.sort_values(by='xPTS', ascending=False).reset_index(drop=True)
        df_show.index = df_show.index + 1
        st.dataframe(pd.DataFrame({
            '순위': df_show.index, '팀명': df_show['Team'], '경기수': df_show['GP'],
            '총 xG': df_show['real_xG'].round(2), '총 xGA': df_show['real_xGA'].round(2),
            '경기당 xG': df_show['avg_xG'].round(2), 'xG 마진': df_show['xGDiff'].round(2),
            '기대 승점': df_show['xPTS'].round(1)
        }), use_container_width=True, hide_index=True)
