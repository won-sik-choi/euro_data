import streamlit as st
import pandas as pd
import numpy as np
import glob
import plotly.express as px
import plotly.graph_objects as go
import os
import asyncio
from scipy.stats import poisson

# 새로운 understat 비동기 라이브러리 불러오기
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
    div[data-testid="stRadio"] > div {
        flex-direction: row;
        flex-wrap: wrap;
        gap: 10px;
    }
    div[data-testid="stRadio"] label {
        background-color: #1b2e4b !important;
        border: 1px solid #3b3f5c !important;
        border-radius: 8px !important;
        padding: 10px 20px !important;
        color: #bfc9d4 !important;
        font-weight: bold !important;
        cursor: pointer;
        width: 23% !important;
        text-align: center;
        margin: 0 !important;
    }
    div[data-testid="stRadio"] label:hover {
        border-color: #2196f3 !important;
        color: #ffffff !important;
    }
    div[data-testid="stRadio"] label[data-checked="true"] {
        background-color: #2196f3 !important;
        border-color: #2196f3 !important;
        color: #ffffff !important;
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
# 2. 데이터 로드 및 전처리 (22개 전 리그 매핑)
# ==========================================
@st.cache_data
def load_data():
    files = sorted(glob.glob("*.xlsx"))
    if not files:
        return {}
    
    league_sheets = {
        # 🏴󠁧󠁢󠁥󠁮󠁧󠁿 잉글랜드
        'EPL (잉글랜드 1부)': 'E0', '챔피언십 (잉글랜드 2부)': 'E1', '리그 원 (잉글랜드 3부)': 'E2', '리그 투 (잉글랜드 4부)': 'E3', '내셔널리그 (잉글랜드 5부)': 'EC',
        # 🏴󠁧󠁢󠁳󠁣󠁴󠁿 스코틀랜드
        '스코티시 프리미어쉽 (스코틀랜드 1부)': 'SC0', '스코티시 챔피언십 (스코틀랜드 2부)': 'SC1', '스코티시 리그 1 (스코틀랜드 3부)': 'SC2', '스코티시 리그 2 (스코틀랜드 4부)': 'SC3',
        # 🇩🇪 독일
        '분데스리가 (독일 1부)': 'D1', '분데스리가 2 (독일 2부)': 'D2',
        # 🇪🇸 스페인
        '라리가 (스페인 1부)': 'SP1', '세군다 디비시온 (스페인 2부)': 'SP2',
        # 🇮🇹 이탈리아
        '세리에 A (이탈리아 1부)': 'I1', '세리에 B (이탈리아 2부)': 'I2',
        # 🇫🇷 프랑스
        '리그 앙 (프랑스 1부)': 'F1', '리그 두 (프랑스 2부)': 'F2',
        # 🇳🇱 🇧🇪 🇵🇹 🇹🇷 🇬🇷 기타 주요 리그
        '에레디비시 (네덜란드 1부)': 'N1', '주필러 프로 리그 (벨기에 1부)': 'B1', '프리메이라리가 (포르투갈 1부)': 'P1', '쉬페르리그 (터키 1부)': 'T1', '수페르리가 엘라다 (그리스 1부)': 'G1'
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
                'Team': t['title'], 'GP': gp,
                'real_xG': total_xg, 'real_xGA': total_xga,
                'avg_xG': total_xg / (gp + 1e-5), 'avg_xGA': total_xga / (gp + 1e-5),
                'xPTS': total_xpts
            })
        return pd.DataFrame(parsed)

@st.cache_data(ttl=3600)
def fetch_understat_xg(selected_league_name):
    if not HAS_UNDERSTAT: return None
    understat_map = {'EPL (잉글랜드 1부)': 'epl', '라리가 (스페인 1부)': 'la_liga', '분데스리가 (독일 1부)': 'bundesliga', '세리에 A (이탈리아 1부)': 'serie_a', '리그 앙 (프랑스 1부)': 'ligue_1'}
    if selected_league_name not in understat_map: return None
    
    league_code = understat_map[selected_league_name]
    for year in [2025, 2024]:
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            df_res = loop.run_until_complete(_async_get_understat(league_code, year))
            loop.close()
            if df_res is not None and not df_res.empty: return df_res
        except Exception: pass
    return None

league_dict = load_data()

st.title("⚽ 선택 매치업 종합 배팅 분석 대시보드")
st.caption("AI 승률/언오버/카드 예측, 95% 환급률 반영 적정 배당 산출 및 경기 흐름 분석 종합 시스템입니다.")

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
# 4. 포아송 및 95% 환급률 배당 계산 함수
# ==========================================
def calculate_poisson_probabilities(df_league, home_team, away_team, h_inj_att=0, h_inj_def=0, a_inj_att=0, a_inj_def=0, xg_df=None):
    avg_home_goals = df_league['FTHG'].mean()
    avg_away_goals = df_league['FTAG'].mean()
    
    home_matches = df_league[df_league['HomeTeam'] == home_team]
    away_matches = df_league[df_league['AwayTeam'] == away_team]
    
    if home_matches.empty or away_matches.empty: return None
    
    home_attack = home_matches['FTHG'].mean() / (avg_home_goals + 1e-5)
    home_defense = home_matches['FTAG'].mean() / (avg_away_goals + 1e-5)
    away_attack = away_matches['FTAG'].mean() / (avg_away_goals + 1e-5)
    away_defense = away_matches['FTHG'].mean() / (avg_home_goals + 1e-5)
    
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

    home_attack = max(0.1, home_attack * (1.0 - h_inj_att * 0.08))
    home_defense = home_defense * (1.0 + h_inj_def * 0.08)
    away_attack = max(0.1, away_attack * (1.0 - a_inj_att * 0.08))
    away_defense = away_defense * (1.0 + a_inj_def * 0.08)
    
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
    prob_under25 = 1.0 - prob_over25
    
    h_cards = (home_matches['HY'].mean() if 'HY' in home_matches.columns else 0)
    a_cards = (away_matches['AY'].mean() if 'AY' in away_matches.columns else 0)
    h_sot = (home_matches['HST'].mean() if 'HST' in home_matches.columns else 0)
    a_sot = (away_matches['AST'].mean() if 'AST' in away_matches.columns else 0)
    h_cor = (home_matches['HC'].mean() if 'HC' in home_matches.columns else 0)
    a_cor = (away_matches['AC'].mean() if 'AC' in away_matches.columns else 0)
    
    # 💡 95% 환급률 반영 적정 배당 계산 (Payout Ratio = 0.95)
    PAYOUT_RATIO = 0.95
    fair_h_odds = PAYOUT_RATIO / (prob_home_win + 1e-5)
    fair_d_odds = PAYOUT_RATIO / (prob_draw + 1e-5)
    fair_a_odds = PAYOUT_RATIO / (prob_away_win + 1e-5)
    
    fair_over_odds = PAYOUT_RATIO / (prob_over25 + 1e-5)
    fair_under_odds = PAYOUT_RATIO / (prob_under25 + 1e-5)
    
    return {
        'h_attack': home_attack, 'h_defense': home_defense,
        'a_attack': away_attack, 'a_defense': away_defense,
        'exp_h_goals': expected_home_goals, 'exp_a_goals': expected_away_goals,
        'home_win': prob_home_win * 100, 'draw': prob_draw * 100, 'away_win': prob_away_win * 100,
        'over25': prob_over25 * 100, 'under25': prob_under25 * 100,
        'fair_home_odds': fair_h_odds,
        'fair_draw_odds': fair_d_odds,
        'fair_away_odds': fair_a_odds,
        'fair_over_odds': fair_over_odds,
        'fair_under_odds': fair_under_odds,
        'h_avg_cards': h_cards, 'a_avg_cards': a_cards,
        'exp_cards': h_cards + a_cards,
        'exp_sot': h_sot + a_sot,
        'exp_corners': h_cor + a_cor
    }

ai_result = calculate_poisson_probabilities(df, home_team, away_team, home_inj_att, home_inj_def, away_inj_att, away_inj_def, df_real_xg)

# ==========================================
# 5. 2줄 카드형 네비게이션 구조
# ==========================================
page_options = [
    "🤖 Page 1: AI 종합 예측 요약",
    "⚔️ Page 2: 맞대결 전적 & 최근 흐름",
    "💰 Page 3: 초기 vs 마감 배당 세부 분석",
    "🟨 Page 4: 매치 심판 판정 성향",
    "📈 Page 5: 홈 vs 원정 유효슈팅 & 지표",
    "📊 Page 6: 리그 실시간 순위표",
    "🎯 Page 7: 팀별 정밀 xG 분석 리포트",
    "🔥 Page 8: AI 경기 흐름 & 베스트 배팅 추천"
]

selected_page = st.radio("📄 **분석 페이지 선택 (2줄 카드 클릭):**", page_options, index=0)
st.markdown("<div style='margin-bottom: 20px;'></div>", unsafe_allow_html=True)

# ------------------------------------------
# Page 1: AI 종합 예측 요약 (95% 환급률 적용)
# ------------------------------------------
if selected_page == page_options[0]:
    st.subheader(f"🤖 AI 종합 예측 리포트 ({home_team} vs {away_team})")
    
    if df_real_xg is not None and not df_real_xg.empty:
        st.success("⚡ **Understat 실제 xG(기대득점) 데이터가 모델에 실시간 반영되었습니다.**")
    else:
        st.caption("포아송 분포, 슈팅 기반 xG 알고리즘 및 결장자 가중치를 종합 반영한 예측 결과입니다.")
    
    if (home_inj_att + home_inj_def + away_inj_att + away_inj_def) > 0:
        st.info(f"🏥 **결장자 설정 반영됨:** {home_team}(공격-{home_inj_att}, 수비-{home_inj_def}) / {away_team}(공격-{away_inj_att}, 수비-{away_inj_def})")
        
    if ai_result:
        st.markdown("##### 🎯 **1. AI 승무패 / 언오버 확률 및 적정 배당 (환급률 95% 기준)**")
        c_ai1, c_ai2, c_ai3, c_ai4, c_ai5 = st.columns(5)
        c_ai1.metric(f"🏠 {home_team} 승리 확률", f"{ai_result['home_win']:.1f}%", f"적정 {ai_result['fair_home_odds']:.2f}")
        c_ai2.metric("🤝 무승부 확률", f"{ai_result['draw']:.1f}%", f"적정 {ai_result['fair_draw_odds']:.2f}")
        c_ai3.metric(f"🚀 {away_team} 승리 확률", f"{ai_result['away_win']:.1f}%", f"적정 {ai_result['fair_away_odds']:.2f}")
        c_ai4.metric("🔥 2.5 오버 확률", f"{ai_result['over25']:.1f}%", f"적정 {ai_result['fair_over_odds']:.2f}")
        c_ai5.metric("🛡️ 2.5 언더 확률", f"{ai_result['under25']:.1f}%", f"적정 {ai_result['fair_under_odds']:.2f}")
        
        st.markdown("---")
        st.markdown("##### 📊 **2. 팀별 보정 전력 지수 & AI 예상 기대득점 (xG)**")
        p1, p2, p3, p4, p5 = st.columns(5)
        p1.metric(f"🏠 {home_team} 홈 공격력", f"{ai_result['h_attack']:.2f}", "1.0 이상 = 우수")
        p2.metric(f"🏠 {home_team} 홈 수비력", f"{ai_result['h_defense']:.2f}", "1.0 이하 = 우수")
        p3.metric(f"🚀 {away_team} 원정 공격력", f"{ai_result['a_attack']:.2f}", "1.0 이상 = 우수")
        p4.metric(f"🚀 {away_team} 원정 수비력", f"{ai_result['a_defense']:.2f}", "1.0 이하 = 우수")
        p5.metric("⚽ AI 예상 스코어 (xG)", f"{ai_result['exp_h_goals']:.2f} : {ai_result['exp_a_goals']:.2f}", f"합계 {ai_result['exp_h_goals']+ai_result['exp_a_goals']:.2f}골")
        
        st.markdown("---")
        st.markdown("##### 🟨 **3. AI 매치 카드(Yellow Cards) 위험도 시뮬레이션**")
        kc1, kc2, kc3 = st.columns(3)
        kc1.metric(f"🏠 {home_team} 홈 경기당 카드 수집", f"{ai_result['h_avg_cards']:.2f} 개")
        kc2.metric(f"🚀 {away_team} 원정 경기당 카드 수집", f"{ai_result['a_avg_cards']:.2f} 개")
        card_risk = "높음 (격렬한 경기 예상)" if ai_result['exp_cards'] >= 4.5 else "보통 (평이한 수준)"
        kc3.metric("🚨 예상 총 카드 수 (Card Market)", f"약 {ai_result['exp_cards']:.1f} 개", f"위험도: {card_risk}")

# ------------------------------------------
# Page 2: 맞대결 전적 (홈/원정 동일 조건 필터 추가)
# ------------------------------------------
elif selected_page == page_options[1]:
    st.subheader(f"⚔️ {home_team} vs {away_team} 맞대결 전적 & 최근 흐름")
    
    # 💡 홈/원정 조건 필터 세렉트박스
    h2h_mode = st.radio("🔍 맞대결 조회 조건 선택:", ["🌐 역대 전체 맞대결", f"🏠 현재 조건 동일 맞대결 ({home_team} 홈 vs {away_team} 원정만)"], horizontal=True)
    
    if h2h_mode.startswith("🌐"):
        h2h = df[((df['HomeTeam'] == home_team) & (df['AwayTeam'] == away_team)) | 
                 ((df['HomeTeam'] == away_team) & (df['AwayTeam'] == home_team))].sort_values(by='Date', ascending=False)
    else:
        h2h = df[(df['HomeTeam'] == home_team) & (df['AwayTeam'] == away_team)].sort_values(by='Date', ascending=False)
        
    st.markdown("##### 📌 **맞대결 요약 (H2H)**")
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
        
        with st.expander("📄 맞대결 경기 스코어 (전반 스코어 포함) 목록 보기"):
            h2h_disp = h2h.copy()
            h2h_disp['Score'] = h2h_disp['FTHG'].astype(int).astype(str) + " : " + h2h_disp['FTAG'].astype(int).astype(str)
            if 'HTHG' in h2h_disp.columns:
                h2h_disp['HT_Score'] = h2h_disp['HTHG'].fillna(0).astype(int).astype(str) + " : " + h2h_disp['HTAG'].fillna(0).astype(int).astype(str)
            else:
                h2h_disp['HT_Score'] = "-"
            
            h2h_disp['Date'] = h2h_disp['Date'].dt.strftime('%Y-%m-%d')
            disp_cols = ['Date', 'Season', 'HomeTeam', 'HT_Score', 'Score', 'AwayTeam']
            renames = {'Date':'날짜', 'Season':'시즌', 'HomeTeam':'홈팀', 'HT_Score':'전반스코어', 'Score':'최종스코어', 'AwayTeam':'원정팀'}
            st.dataframe(h2h_disp[disp_cols].rename(columns=renames), use_container_width=True, hide_index=True)
    else:
        st.info("선택하신 조건에 해당하는 맞대결 기록이 없습니다.")
        
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
elif selected_page == page_options[2]:
    st.subheader(f"💰 초기 배당 vs 마감 배당(Closing Odds) 세부 분석")
    col_in1, col_in2, col_in3, col_in4 = st.columns(4)
    with col_in1: input_h_odds = st.number_input(f"🏠 {home_team} 홈 승리 배당", min_value=1.01, max_value=20.0, value=1.75, step=0.05)
    with col_in2: input_d_odds = st.number_input("🤝 무승부 배당", min_value=1.01, max_value=20.0, value=3.40, step=0.05)
    with col_in3: input_a_odds = st.number_input(f"🚀 {away_team} 원정 승리 배당", min_value=1.01, max_value=20.0, value=4.50, step=0.05)
    with col_in4: tolerance = st.slider("배당 허용 오차 범위 (±)", min_value=0.02, max_value=0.30, value=0.10, step=0.02)
    
    st.markdown("---")
    if 'B365H' in df.columns:
        similar_games = df[(df['B365H'] >= input_h_odds - tolerance) & (df['B365H'] <= input_h_odds + tolerance)].copy()
        if not similar_games.empty:
            tot_cnt = len(similar_games)
            h_win = len(similar_games[similar_games['FTR'] == 'H'])
            draw = len(similar_games[similar_games['FTR'] == 'D'])
            a_win = len(similar_games[similar_games['FTR'] == 'A'])
            
            has_close = ('B365CH' in similar_games.columns) and ('B365CD' in similar_games.columns) and ('B365CA' in similar_games.columns)
            drop_win_rate = 0.0
            if has_close:
                drop_h = similar_games[similar_games['B365CH'] < similar_games['B365H']]
                drop_h_win = len(drop_h[drop_h['FTR'] == 'H']) if len(drop_h) > 0 else 0
                drop_win_rate = (drop_h_win / len(drop_h) * 100) if len(drop_h) > 0 else 0
            
            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.metric("홈 승리 적중률", f"{(h_win/tot_cnt)*100:.1f}%", f"{h_win} / {tot_cnt} 경기")
            mc2.metric("무승부 발생률", f"{(draw/tot_cnt)*100:.1f}%", f"{draw} / {tot_cnt} 경기")
            mc3.metric("원정 승리 발생률", f"{(a_win/tot_cnt)*100:.1f}%", f"{a_win} / {tot_cnt} 경기")
            if has_close: mc4.metric("마감 시 홈배당 하락 시 승률", f"{drop_win_rate:.1f}%")
            
            st.markdown("---")
            similar_games['Score'] = similar_games['FTHG'].astype(int).astype(str) + " : " + similar_games['FTAG'].astype(int).astype(str)
            similar_games['DateStr'] = similar_games['Date'].dt.strftime('%Y-%m-%d')
            show_cols = ['DateStr', 'HomeTeam', 'AwayTeam', 'B365H', 'B365D', 'B365A', 'Score', 'FTR']
            st.dataframe(similar_games[show_cols].rename(columns={'DateStr':'날짜', 'HomeTeam':'홈팀', 'AwayTeam':'원정팀', 'B365H':'초기홈', 'B365D':'초기무', 'B365A':'초기원정', 'Score':'스코어', 'FTR':'결과'}), use_container_width=True, hide_index=True)

# ------------------------------------------
# Page 4: 매치 심판 성향
# ------------------------------------------
elif selected_page == page_options[3]:
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
elif selected_page == page_options[4]:
    st.subheader(f"📈 {home_team} (홈 성적) vs {away_team} (원정 성적) 슈팅 & 코너킥 세부 비교")
    home_only = df[df['HomeTeam'] == home_team]
    away_only = df[df['AwayTeam'] == away_team]
    
    if not home_only.empty and not away_only.empty:
        sc1, sc2, sc3, sc4 = st.columns(4)
        sc1.metric(f"🏠 {home_team} 홈 득점", f"{home_only['FTHG'].mean():.2f} 골")
        sc2.metric(f"🏠 {home_team} 홈 실점", f"{home_only['FTAG'].mean():.2f} 골")
        sc3.metric(f"🚀 {away_team} 원정 득점", f"{away_only['FTAG'].mean():.2f} 골")
        sc4.metric(f"🚀 {away_team} 원정 실점", f"{away_only['FTHG'].mean():.2f} 골")
        
        st.markdown("---")
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
# Page 6: 실시간 순위표
# ------------------------------------------
elif selected_page == page_options[5]:
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
# Page 7: 팀별 정밀 xG 분석
# ------------------------------------------
elif selected_page == page_options[6]:
    st.subheader(f"🎯 {selected_league} 정밀 기대득점(xG) 통계 리포트")
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

# ------------------------------------------
# Page 8: AI 경기 흐름 & 베스트 배팅 조건
# ------------------------------------------
elif selected_page == page_options[7]:
    st.subheader(f"🔥 AI 경기 흐름 분석 & 베스트 배팅 조건 ({home_team} vs {away_team})")
    
    if ai_result:
        st.markdown("##### 🔍 **1. AI 종합 경기 흐름 양상 리포트**")
        flow_desc = []
        if ai_result['exp_h_goals'] + ai_result['exp_a_goals'] >= 3.0:
            flow_desc.append("⚽ **난타전 예상:** 양 팀의 기대 득점(xG) 합계가 높은 활발한 공격 위주의 경기 양상입니다.")
        elif ai_result['exp_h_goals'] + ai_result['exp_a_goals'] <= 2.1:
            flow_desc.append("🛡️ **소강상태 예상:** 수비적 탐색전이 길어지며 1~2골 내외의 팽팽한 소점수 양상이 펼쳐질 가능성이 큽니다.")
        else:
            flow_desc.append("⚖️ **평이한 흐름:** 일반적인 리그 평균 수준의 득점 기회가 창출될 것으로 예상됩니다.")
            
        if ai_result['exp_cards'] >= 4.5:
            flow_desc.append("🚨 **거친 파울 매치:** 주심 판정 성향 및 양 팀 파울 페이스상 많은 경고 카드가 유발될 가능성이 높습니다.")
        if ai_result['exp_sot'] >= 9.5:
            flow_desc.append("🎯 **공격적 템포:** 유효슈팅 창출 빈도가 높아 공격 전개가 매우 빠르게 진행될 경기입니다.")
            
        for desc in flow_desc:
            st.info(desc)
            
        st.markdown("---")
        st.markdown("##### 🏆 **2. 가장 가치(Value) 높은 베스트 배팅 조건 Top 3 (95% 환급률 적용)**")
        
        ranked_picks = []
        if ai_result['home_win'] >= 60.0:
            ranked_picks.append((ai_result['home_win'], "🔥 [승무패] " + home_team + " 홈 승리", f"확률 {ai_result['home_win']:.1f}% (적정 배당 {ai_result['fair_home_odds']:.2f})"))
        elif ai_result['away_win'] >= 55.0:
            ranked_picks.append((ai_result['away_win'], "🔥 [승무패] " + away_team + " 원정 승리", f"확률 {ai_result['away_win']:.1f}% (적정 배당 {ai_result['fair_away_odds']:.2f})"))
            
        if ai_result['over25'] >= 58.0:
            ranked_picks.append((ai_result['over25'], "⚽ [언오버] 2.5 오버 (Over)", f"확률 {ai_result['over25']:.1f}% (적정 배당 {ai_result['fair_over_odds']:.2f})"))
        elif ai_result['under25'] >= 58.0:
            ranked_picks.append((ai_result['under25'], "🛡️ [언오버] 2.5 언더 (Under)", f"확률 {ai_result['under25']:.1f}% (적정 배당 {ai_result['fair_under_odds']:.2f})"))
            
        if ai_result['exp_cards'] >= 4.5:
            ranked_picks.append((70.0, "🟨 [카드 마켓] 오버 4.5 카드가능성", f"예상 총 카드 수 약 {ai_result['exp_cards']:.1f}개"))
        if ai_result['exp_sot'] >= 9.5:
            ranked_picks.append((65.0, "🎯 [유효슈팅] 오버 9.5 유효슈팅", f"양 팀 예상 합계 {ai_result['exp_sot']:.1f}개"))

        ranked_picks.sort(key=lambda x: x[0], reverse=True)
        top_picks = ranked_picks[:3]
        
        if top_picks:
            b_cols = st.columns(len(top_picks))
            for idx, (score, title, sub) in enumerate(top_picks):
                with b_cols[idx]:
                    st.markdown(f"""
                    <div style='background-color:#1b2e4b; border:2px solid #009688; border-radius:12px; padding:15px;'>
                        <p style='color:#009688; margin:0;'><b>BEST PICK #{idx+1}</b></p>
                        <h3 style='color:#ffffff; margin:8px 0;'><b>{title}</b></h3>
                        <p style='color:#bfc9d4; margin:0; font-size:0.9rem;'>{sub}</p>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("양 팀 전력이 팽팽하여 전반전 경기 진행을 확인한 후 라이브 배팅으로 접근하시는 것을 권장합니다.")
            
        st.markdown("---")
        st.markdown("##### 📊 **3. 항목별 종합 AI 예측 수치 표**")
        summary_data = {
            '분석 항목': ['승무패 적정 배당 (95%)', '2.5 골 언오버 적정 배당', '예상 기대 득점 (xG)', '예상 총 카드 수', '예상 총 유효슈팅 수', '예상 총 코너킥 수'],
            'AI 예상 분석 수치': [
                f"홈승 {ai_result['fair_home_odds']:.2f} / 무승부 {ai_result['fair_draw_odds']:.2f} / 원정승 {ai_result['fair_away_odds']:.2f}",
                f"2.5 오버 배당 {ai_result['fair_over_odds']:.2f} / 2.5 언더 배당 {ai_result['fair_under_odds']:.2f}",
                f"{home_team} {ai_result['exp_h_goals']:.2f} 골 : {away_team} {ai_result['exp_a_goals']:.2f} 골 (합계 {ai_result['exp_h_goals']+ai_result['exp_a_goals']:.2f}골)",
                f"약 {ai_result['exp_cards']:.1f} 개 (홈 {ai_result['h_avg_cards']:.1f} / 원정 {ai_result['a_avg_cards']:.1f})",
                f"약 {ai_result['exp_sot']:.1f} 개",
                f"약 {ai_result['exp_corners']:.1f} 개"
            ]
        }
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)
