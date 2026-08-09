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

# ------------------------------------------
# 💡 이적 데이터셋 로더 & 팀명 표준화 개선
# ------------------------------------------
@st.cache_data(ttl=86400)
def load_github_transfer_dataset():
    url = "https://raw.githubusercontent.com/davidcariboo/transfermarkt-datasets/master/data/transfers.csv"
    try:
        df_trans = pd.read_csv(url)
        return df_trans
    except Exception:
        return None

def fetch_real_transfers(team_name, full_trans_df=None):
    # 주요 클럽 키워드 매칭 보정
    clean_team = team_name.lower().replace(" FC", "").replace(" fc", "").strip()
    
    if full_trans_df is not None and not full_trans_df.empty:
        # 최근 시즌 데이터로 정렬
        if 'transfer_season' in full_trans_df.columns:
            full_trans_df = full_trans_df.sort_values(by='transfer_season', ascending=False)
            
        in_mask = full_trans_df['to_club_name'].astype(str).str.lower().str.contains(clean_team, na=False)
        out_mask = full_trans_df['from_club_name'].astype(str).str.lower().str.contains(clean_team, na=False)
        
        df_in = full_trans_df[in_mask].head(5).copy()
        df_out = full_trans_df[out_mask].head(5).copy()
        
        records = []
        for _, r in df_in.iterrows():
            fee_raw = r.get('transfer_fee', 0)
            fee_val = float(fee_raw) / 1e6 if pd.notna(fee_raw) and str(fee_raw).replace('.','',1).isdigit() else 0.0
            records.append({
                '구분': '영입 (IN)',
                '선수명': r.get('player_name', '알수없음'),
                '전/후 클럽': r.get('from_club_name', '-'),
                '이적료 (€M)': round(fee_val, 1),
                '전력 영향도': f"+{round(fee_val * 0.1, 1)}%"
            })
            
        for _, r in df_out.iterrows():
            fee_raw = r.get('transfer_fee', 0)
            fee_val = float(fee_raw) / 1e6 if pd.notna(fee_raw) and str(fee_raw).replace('.','',1).isdigit() else 0.0
            records.append({
                '구분': '방출 (OUT)',
                '선수명': r.get('player_name', '알수없음'),
                '전/후 클럽': r.get('to_club_name', '-'),
                '이적료 (€M)': round(fee_val, 1),
                '전력 영향도': f"-{round(fee_val * 0.08, 1)}%"
            })
            
        if records:
            df_res = pd.DataFrame(records)
            in_fee = df_res[df_res['구분'] == '영입 (IN)']['이적료 (€M)'].sum()
            out_fee = df_res[df_res['구분'] == '방출 (OUT)']['이적료 (€M)'].sum()
            net_spend = in_fee - out_fee
            power_change = round((in_fee * 0.1) - (out_fee * 0.08), 2)
            
            return {
                'df': df_res,
                'in_fee': in_fee,
                'out_fee': out_fee,
                'net_spend': net_spend,
                'power_change_pct': power_change
            }
            
    empty_df = pd.DataFrame(columns=['구분', '선수명', '전/후 클럽', '이적료 (€M)', '전력 영향도'])
    return {
        'df': empty_df,
        'in_fee': 0.0,
        'out_fee': 0.0,
        'net_spend': 0.0,
        'power_change_pct': 0.0
    }

league_dict = load_data()

st.title("⚽ 선택 매치업 종합 배팅 분석 대시보드")
st.caption("AI 승률/언오버/카드 예측 및 전/후반 득점, 마감 배당, 실시간 xG 수집, 실제 이적 전력 변화까지 정밀 분석이 가능한 종합 시스템입니다.")

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

# 실제 오픈소스 이적 데이터셋 연결
full_trans_df = load_github_transfer_dataset()
home_trans = fetch_real_transfers(home_team, full_trans_df)
away_trans = fetch_real_transfers(away_team, full_trans_df)

# ==========================================
# 4. 포아송, 결장자, xG 및 이적 보정 계산 함수
# ==========================================
def calculate_poisson_probabilities(df_league, home_team, away_team, h_inj_att=0, h_inj_def=0, a_inj_att=0, a_inj_def=0, xg_df=None, h_trans_p=0, a_trans_p=0):
    avg_home_goals = df_league['FTHG'].mean()
    avg_away_goals = df_league['FTAG'].mean()
    
    home_matches = df_league[df_league['HomeTeam'] == home_team]
    away_matches = df_league[df_league['AwayTeam'] == away_team]
    
    if home_matches.empty or away_matches.empty:
        return None
    
    # 1. 기본 공격력/수비력 지수 계산
    home_attack = home_matches['FTHG'].mean() / (avg_home_goals + 1e-5)
    home_defense = home_matches['FTAG'].mean() / (avg_away_goals + 1e-5)
    
    away_attack = away_matches['FTAG'].mean() / (avg_away_goals + 1e-5)
    away_defense = away_matches['FTHG'].mean() / (avg_home_goals + 1e-5)
    
    # 2. 🔄 이적시장 전력 변동률 반영
    home_attack = home_attack * (1.0 + (h_trans_p / 100.0))
    away_attack = away_attack * (1.0 + (a_trans_p / 100.0))
    
    # 3. 🎯 xG 보정 (5대 리그 실제 xG 우선, 없으면 슈팅 기반 보정)
    if xg_df is not None and not xg_df.empty:
        h_row = xg_df[xg_df['Team'].str.contains(home_team[:4], case=False, na=False)]
        a_row = xg_df[xg_df['Team'].str.contains(away_team[:4], case=False, na=False)]
        
        if not h_row.empty and not a_row.empty:
            avg_league_xg = xg_df['avg_xG'].mean() + 1e-5
            home_attack = (h_row['avg_xG'].values[0] / avg_league_xg) * (1.0 + (h_trans_p / 100.0))
            away_attack = (a_row['avg_xG'].values[0] / avg_league_xg) * (1.0 + (a_trans_p / 100.0))
            home_defense = (h_row['avg_xGA'].values[0] / avg_league_xg)
            away_defense = (a_row['avg_xGA'].values[0] / avg_league_xg)
            
    elif 'HST' in home_matches.columns and 'AST' in away_matches.columns and not home_matches['HST'].isna().all():
        h_xg_stat = (home_matches['HST'].mean() * 0.32) + ((home_matches['HS'].mean() - home_matches['HST'].mean()) * 0.06)
        a_xg_stat = (away_matches['AST'].mean() * 0.32) + ((away_matches['AS'].mean() - away_matches['AST'].mean()) * 0.06)
        
        league_h_xg = (df_league['HST'].mean() * 0.32) + ((df_league['HS'].mean() - df_league['HST'].mean()) * 0.06)
        league_a_xg = (df_league['AST'].mean() * 0.32) + ((df_league['AS'].mean() - df_league['AST'].mean()) * 0.06)
        
        home_attack = (home_attack * 0.5) + ((h_xg_stat / (league_h_xg + 1e-5)) * 0.5)
        away_attack = (away_attack * 0.5) + ((a_xg_stat / (league_a_xg + 1e-5)) * 0.5)

    # 4. 🏥 결장자 가중치 반영 (인당 8% 조정)
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
    
    prob_over25 = 0.0
    for i in range(max_goals):
        for j in range(max_goals):
            if i + j > 2.5:
                prob_over25 += matrix[i, j]
                
    # 카드 지표 계산
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

ai_result = calculate_poisson_probabilities(
    df, home_team, away_team, 
    home_inj_att, home_inj_def, away_inj_att, away_inj_def, 
    df_real_xg, home_trans['power_change_pct'], away_trans['power_change_pct']
)

# ==========================================
# 5. 8개 페이지 구성
# ==========================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "🤖 Page 1: AI 종합 예측 & 카드 시뮬레이션",
    "⚔️ Page 2: 맞대결 전적 & 전/후반 득점",
    "💰 Page 3: 초기 vs 마감 배당 세부 분석",
    "🟨 Page 4: 매치 심판 상세 판정 성향",
    "📈 Page 5: 홈 vs 원정 유효슈팅 & 지표",
    "📊 Page 6: 리그 실시간 순위표",
    "🎯 Page 7: 팀별 정밀 xG 분석 리포트",
    "🔄 Page 8: 이적 현황 & 전력 변화 분석"
])

# ------------------------------------------
# Page 1: AI 종합 예측 & 카드 시뮬레이션
# ------------------------------------------
with tab1:
    st.subheader(f"🤖 AI 종합 예측 리포트 ({home_team} vs {away_team})")
    
    if df_real_xg is not None and not df_real_xg.empty:
        st.success("⚡ **Understat 실제 xG 및 이적 시장 보정 수치가 AI 모델에 실시간 반영되었습니다.**")
    else:
        st.caption("포아송 분포, 슈팅 기반 xG 알고리즘, 이적 시장 전력 변동 및 결장자를 종합 반영한 예측 결과입니다.")
    
    if (home_inj_att + home_inj_def + away_inj_att + away_inj_def) > 0:
        st.info(f"🏥 **결장자 설정 반영됨:** {home_team}(공격-{home_inj_att}, 수비-{home_inj_def}) / {away_team}(공격-{away_inj_att}, 수비-{away_inj_def})")
        
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
        st.markdown("##### 📊 **2. 팀별 보정 전력 지수 & AI 예상 기대득점 (xG)**")
        p1, p2, p3, p4, p5 = st.columns(5)
        p1.metric(f"🏠 {home_team} 홈 공격력", f"{ai_result['h_attack']:.2f}", "1.0 이상 = 우수")
        p2.metric(f"🏠 {home_team} 홈 수비력", f"{ai_result['h_defense']:.2f}", "1.0 이하 = 우수")
        p3.metric(f"🚀 {away_team} 원정 공격력", f"{ai_result['a_attack']:.2f}", "1.0 이상 = 우수")
        p4.metric(f"🚀 {away_team} 원정 수비력", f"{ai_result['a_defense']:.2f}", "1.0 이하 = 우수")
        p5.metric("⚽ AI 예상 스코어 (xG)", f"{ai_result['exp_h_goals']:.2f} : {ai_result['exp_a_goals']:.2f}", f"합계 {ai_result['exp_h_goals']+ai_result['exp_a_goals']:.2f}골")
        
        st.markdown("---")
        
        # 3. AI 카드(경고/퇴장) 시뮬레이션
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
        st.info("💡 이 리그 데이터에는 주심(Referee) 정보가 포함되어 있지 않습니다.")

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

# ------------------------------------------
# Page 6: 리그 실시간 순위표
# ------------------------------------------
with tab6:
    st.subheader(f"📊 {selected_league} 실시간 리그 순위표")
    st.caption("업로드된 경기 결과를 바탕으로 경기수, 승무패, 득/실점, 득실차, 승점을 자동 계산합니다.")
    
    latest_season = df['Season'].max() if 'Season' in df.columns else None
    df_season = df[df['Season'] == latest_season] if latest_season else df
    
    standings = {t: {'GP': 0, 'W': 0, 'D': 0, 'L': 0, 'GF': 0, 'GA': 0, 'PTS': 0} for t in teams}
    
    for _, r in df_season.iterrows():
        ht, at = r['HomeTeam'], r['AwayTeam']
        hg, ag = r['FTHG'], r['FTAG']
        ftr = r['FTR']
        
        if ht in standings and at in standings and pd.notna(hg) and pd.notna(ag):
            standings[ht]['GP'] += 1
            standings[ht]['GF'] += int(hg)
            standings[ht]['GA'] += int(ag)
            
            standings[at]['GP'] += 1
            standings[at]['GF'] += int(ag)
            standings[at]['GA'] += int(hg)
            
            if ftr == 'H':
                standings[ht]['W'] += 1
                standings[ht]['PTS'] += 3
                standings[at]['L'] += 1
            elif ftr == 'A':
                standings[at]['W'] += 1
                standings[at]['PTS'] += 3
                standings[ht]['L'] += 1
            elif ftr == 'D':
                standings[ht]['D'] += 1
                standings[ht]['PTS'] += 1
                standings[at]['D'] += 1
                standings[at]['PTS'] += 1

    df_rank = pd.DataFrame.from_dict(standings, orient='index').reset_index()
    df_rank.rename(columns={'index': '팀명'}, inplace=True)
    df_rank['GD'] = df_rank['GF'] - df_rank['GA']
    
    df_rank = df_rank.sort_values(by=['PTS', 'GD', 'GF'], ascending=[False, False, False]).reset_index(drop=True)
    df_rank.index = df_rank.index + 1
    df_rank['순위'] = df_rank.index
    
    def highlight_matchup(team_name):
        if team_name == home_team:
            return f"🏠 {team_name} [홈]"
        elif team_name == away_team:
            return f"🚀 {team_name} [원정]"
        return team_name

    df_rank['팀명'] = df_rank['팀명'].apply(highlight_matchup)
    
    df_rank_display = df_rank[['순위', '팀명', 'GP', 'W', 'D', 'L', 'GF', 'GA', 'GD', 'PTS']]
    df_rank_display.columns = ['순위', '팀명', '경기수(GP)', '승(W)', '무(D)', '패(L)', '득점(GF)', '실점(GA)', '득실차(GD)', '승점(PTS)']
    
    st.dataframe(
        df_rank_display,
        use_container_width=True,
        hide_index=True
    )

# ------------------------------------------
# Page 7: 팀별 정밀 xG 분석 리포트
# ------------------------------------------
with tab7:
    st.subheader(f"🎯 {selected_league} 정밀 기대득점(xG) 통계 리포트")
    
    if df_real_xg is not None and not df_real_xg.empty:
        st.success("⚡ **Understat 라이브 xG 데이터를 수집하여 표시 중입니다.**")
        df_real_xg_disp = df_real_xg.copy()
        df_real_xg_disp['xGDiff'] = df_real_xg_disp['real_xG'] - df_real_xg_disp['real_xGA']
        df_real_xg_disp = df_real_xg_disp.sort_values(by='xPTS', ascending=False).reset_index(drop=True)
        df_real_xg_disp.index = df_real_xg_disp.index + 1
        
        df_show = pd.DataFrame({
            '순위': df_real_xg_disp.index,
            '팀명': df_real_xg_disp['Team'],
            '경기수': df_real_xg_disp['GP'],
            '총 기대득점 (xG)': df_real_xg_disp['real_xG'].round(2),
            '총 기대실점 (xGA)': df_real_xg_disp['real_xGA'].round(2),
            '경기당 xG': df_real_xg_disp['avg_xG'].round(2),
            '경기당 xGA': df_real_xg_disp['avg_xGA'].round(2),
            'xG 마진 (xG - xGA)': df_real_xg_disp['xGDiff'].round(2),
            '기대 승점 (xPTS)': df_real_xg_disp['xPTS'].round(1)
        })
        st.dataframe(df_show, use_container_width=True, hide_index=True)
    else:
        st.info("💡 Understat 라이브 연동 미작동 시, **엑셀 슈팅/유효슈팅 알고리즘 기반 xG 분석표**를 자동 생성합니다.")
        latest_season = df['Season'].max() if 'Season' in df.columns else None
        df_season = df[df['Season'] == latest_season] if latest_season else df
        
        xg_summary = []
        for t in teams:
            h_m = df_season[df_season['HomeTeam'] == t]
            a_m = df_season[df_season['AwayTeam'] == t]
            gp = len(h_m) + len(a_m)
            
            if gp > 0:
                h_hst = h_m['HST'].mean() if 'HST' in h_m.columns else 0
                h_hs = h_m['HS'].mean() if 'HS' in h_m.columns else 0
                a_ast = a_m['AST'].mean() if 'AST' in a_m.columns else 0
                a_as = a_m['AS'].mean() if 'AS' in a_m.columns else 0
                
                calc_xg = ((h_hst * 0.32) + ((h_hs - h_hst) * 0.06)) * len(h_m) + ((a_ast * 0.32) + ((a_as - a_ast) * 0.06)) * len(a_m)
                avg_xg = calc_xg / gp
                
                xg_summary.append({
                    '팀명': t,
                    '경기수': gp,
                    '추정 총 xG': round(calc_xg, 2),
                    '경기당 평균 xG': round(avg_xg, 2)
                })
                
        if xg_summary:
            df_calc_xg = pd.DataFrame(xg_summary).sort_values(by='추정 총 xG', ascending=False).reset_index(drop=True)
            df_calc_xg.index = df_calc_xg.index + 1
            df_calc_xg['순위'] = df_calc_xg.index
            st.dataframe(df_calc_xg[['순위', '팀명', '경기수', '추정 총 xG', '경기당 평균 xG']], use_container_width=True, hide_index=True)

# ------------------------------------------
# Page 8: 이적 현황 & 전력 변화 분석 (자동 데이터셋 연동)
# ------------------------------------------
with tab8:
    st.subheader(f"🔄 매치업 팀별 오픈소스 이적 현황 및 순 전력 변화 분석")
    st.caption("GitHub 데이터셋으로부터 매치업 팀의 실제 영입/방출 선수명, 이적료 및 전력 보정치(+/- %)를 수집하여 대조합니다.")
    
    col_t1, col_t2 = st.columns(2)
    
    with col_t1:
        st.markdown(f"##### 🏠 **{home_team} 이적 요약**")
        st.metric(
            "영입/방출 총 지출 (Net Spend)", 
            f"€{home_trans['net_spend']:.1f}M", 
            f"전력 변동률: {home_trans['power_change_pct']:+}%"
        )
        if not home_trans['df'].empty:
            st.dataframe(home_trans['df'], use_container_width=True, hide_index=True)
        else:
            st.info(f"{home_team}의 최근 등록된 이적 내역이 없습니다.")
        
    with col_t2:
        st.markdown(f"##### 🚀 **{away_team} 이적 요약**")
        st.metric(
            "영입/방출 총 지출 (Net Spend)", 
            f"€{away_trans['net_spend']:.1f}M", 
            f"전력 변동률: {away_trans['power_change_pct']:+}%"
        )
        if not away_trans['df'].empty:
            st.dataframe(away_trans['df'], use_container_width=True, hide_index=True)
        else:
            st.info(f"{away_team}의 최근 등록된 이적 내역이 없습니다.")
        
    st.markdown("---")
    st.info("💡 **AI 모델 반영 방식:** 이적 데이터베이스 기반 총 지출 수치 및 전력 변동률이 **Page 1의 공격력/수비력 지수와 AI 예상 승률**에 자동으로 가중 반영됩니다.")
