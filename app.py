"""
국민연금 사업장 영업 대시보드
기존 HTML 대시보드 디자인 그대로 + enrich 내장
실행: streamlit run app.py
"""
import io, re, time, random, requests
import pandas as pd
import numpy as np
import streamlit as st
import folium
from streamlit_folium import st_folium
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="국민연금 가입 사업장 내역", page_icon="🏢",
                   layout="wide", initial_sidebar_state="collapsed")

# ── CSS ─────────────────────────────────────────────────
st.markdown("""<style>
#MainMenu,footer,header,[data-testid="stToolbar"]{visibility:hidden!important}
[data-testid="stSidebar"]{display:none!important}
.block-container{padding:1rem 1rem 0 1rem!important;max-width:100%!important}
.stat-card{background:#f8f9fa;border-radius:8px;padding:10px 12px;border:1px solid #dee2e6;text-align:center}
.stat-num{font-size:20px;font-weight:700;color:#3b82f6}
.stat-lbl{font-size:10px;color:#868e96;margin-top:2px}
.list-item-html{padding:8px 0;border-bottom:1px solid #f1f3f5}
.item-name{font-size:12px;font-weight:600;color:#212529}
.item-sub{font-size:10px;color:#868e96}
.pill-c{background:#dbeafe;color:#1d4ed8;padding:1px 6px;border-radius:3px;font-size:9px;font-weight:600;margin-right:3px}
.pill-i{background:#d1fae5;color:#065f46;padding:1px 6px;border-radius:3px;font-size:9px;font-weight:600;margin-right:3px}
div.stButton > button{background:#3b82f6!important;color:#fff!important;border:none!important;border-radius:6px!important;font-weight:600!important;width:100%!important}
div.stButton > button:hover{background:#2563eb!important}
div.stDownloadButton > button{background:#10b981!important;color:#fff!important;border:none!important;border-radius:6px!important;font-weight:600!important;width:100%!important}
</style>""", unsafe_allow_html=True)

# ── API 키 ─────────────────────────────────────────────
FSC_KEY   = "2446da601452d231f1ff77a4c5ac9598896028fb96e8be586bf7f4be58b6231c"
KAKAO_KEY = "e8eef7246a2cd0f3d2d55af8a806e3ef"
NAV_HDR   = {
    "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/604.1",
    "Referer": "https://map.naver.com/", "Accept-Language": "ko-KR,ko;q=0.9"
}

# ── API 함수 ────────────────────────────────────────────
def fsc(name, biz_no=""):
    try:
        r = requests.get(
            "https://apis.data.go.kr/1160100/service/GetCorpBasicInfoService_V2/getCorpOutline_V2",
            params={"serviceKey":FSC_KEY,"pageNo":1,"numOfRows":5,"resultType":"json","corpNm":name},
            timeout=5)
        items = r.json().get("response",{}).get("body",{}).get("items",{}).get("item",[])
        if not items: return None
        if isinstance(items, dict): items = [items]
        bc = re.sub(r'\D','',str(biz_no))
        m = next((i for i in items if re.sub(r'\D','',str(i.get("bzno",""))) == bc), items[0])
        return {"phone": m.get("enpTlno",""), "addr": m.get("enpBsadr","") or m.get("enpDtadr","")}
    except: return None

def kakao_kw(name, addr):
    try:
        r = requests.get("https://dapi.kakao.com/v2/local/search/keyword.json",
            headers={"Authorization":f"KakaoAK {KAKAO_KEY}"},
            params={"query":f"{name} {addr[:20]}","size":5}, timeout=5)
        docs = r.json().get("documents",[])
        m = next((d for d in docs if name[:4] in d.get("place_name","")), None)
        return m or (docs[0] if docs else None)
    except: return None

def naver_phone(name, addr):
    clean = re.sub(r'주식회사|유한회사|\(주\)|\(유\)|[(){}\[\]]','',name).strip()
    gu = re.search(r'[\w]+구', addr); gu = gu.group() if gu else ""
    try:
        r = requests.get("https://m.map.naver.com/search2/search.nhn",
            headers=NAV_HDR, params={"query":f"{clean} {gu}","type":"all"}, timeout=8)
        if r.status_code == 200:
            for pat in [r'tel:(0\d{1,2}-?\d{3,4}-?\d{4})',r'(?<!\d)(0\d{1,2}-\d{3,4}-\d{4})(?!\d)']:
                f = re.findall(pat, r.text)
                if f: return f[0]
    except: pass
    return ""

def make_excel(df):
    wb = Workbook(); ws = wb.active; ws.title = "영업DB"
    def thin():
        s = Side(style="thin",color="D0D0D0")
        return Border(left=s,right=s,top=s,bottom=s)
    hdrs = list(df.columns)
    widths = {"사업장명":26,"업종":20,"법인여부":8,"주소":44,"전화번호":14,"가입자수":8,"당월고지금액":14,"위도":12,"경도":12}
    CENTER = {"법인여부","가입자수","당월고지금액","위도","경도","전화번호"}
    for c,h in enumerate(hdrs,1):
        cell = ws.cell(1,c,h)
        cell.fill = PatternFill("solid",start_color="1F3864")
        cell.font = Font(name="Arial",size=10,bold=True,color="FFFFFF")
        cell.alignment = Alignment(horizontal="center",vertical="center")
        cell.border = thin()
        ws.column_dimensions[get_column_letter(c)].width = widths.get(h,12)
    ws.row_dimensions[1].height = 18
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(hdrs))}1"
    for ri,row in enumerate(df.itertuples(index=False),2):
        for ci,(h,v) in enumerate(zip(hdrs,row),1):
            if pd.isna(v) or str(v) in ("nan","None"): v=""
            cell = ws.cell(ri,ci,v)
            cell.font = Font(name="Arial",size=10)
            cell.border = thin()
            cell.fill = PatternFill("solid",start_color="FFFFFF" if ri%2==0 else "F8FAFC")
            cell.alignment = Alignment(horizontal="center" if h in CENTER else "left",vertical="center")
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

st.markdown("## 🏢 국민연금 가입 사업장 내역")

# ── 메인 레이아웃: 사이드바 + 지도 ─────────────────────
col_side, col_main = st.columns([1.2, 5], gap="small")

with col_side:
    # 파일 업로드
    uploaded = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")

    if uploaded:
        df_all = pd.read_excel(uploaded)
        df_all.columns = [c.strip() for c in df_all.columns]

        addr_col = next((c for c in df_all.columns if "주소" in c or "도로명" in c), None)
        name_col = next((c for c in df_all.columns if "사업장명" in c or "상호" in c), None)
        type_col = next((c for c in df_all.columns if "형태" in c or "법인" in c), None)
        emp_col  = next((c for c in df_all.columns if "가입자" in c), None)
        pay_col  = next((c for c in df_all.columns if "고지" in c), None)
        biz_col  = next((c for c in df_all.columns if "사업자" in c and "번호" in c), None)
        ind_col  = next((c for c in df_all.columns if "업종" in c), None)

        # 시도/구군 추출
        if addr_col:
            df_all["_시도"] = df_all[addr_col].str.extract(r'^([\S]+시|[\S]+도)')
            df_all["_구군"] = df_all[addr_col].str.extract(r'([\S]+구|[\S]+군)')

        # 시도 선택
        sidos = ["선택"] + sorted(df_all["_시도"].dropna().unique().tolist())
        sido  = st.selectbox("", sidos, label_visibility="collapsed", key="sido")

        df_sido = df_all[df_all["_시도"]==sido] if sido != "선택" else pd.DataFrame()

        # 구/군 선택
        guguns = ["선택"] + (sorted(df_sido["_구군"].dropna().unique().tolist()) if not df_sido.empty else [])
        gu     = st.selectbox("", guguns, label_visibility="collapsed", key="gu")

        df_cur = df_sido[df_sido["_구군"]==gu].copy() if gu != "선택" else pd.DataFrame()
        if emp_col and not df_cur.empty:
            df_cur = df_cur.sort_values(emp_col, ascending=False)

        # 통계
        total = len(df_cur)
        corp  = int((df_cur[type_col].astype(str).str.contains("법인|^1$")).sum()) if type_col and not df_cur.empty else 0
        emp   = int(df_cur[emp_col].sum()) if emp_col and not df_cur.empty else 0
        avg_e = int(df_cur[emp_col].mean()) if emp_col and total > 0 else 0

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("전체",f"{total:,}개")
        c2.metric("법인",f"{corp:,}개")
        c3.metric("총직원",f"{emp:,}명")
        c4.metric("평균",f"{avg_e:,}명")

        # 검색
        search = st.text_input("", placeholder="사업장명 검색...", label_visibility="collapsed")
        if search and name_col and not df_cur.empty:
            df_cur = df_cur[df_cur[name_col].str.contains(search, na=False)]

        # 리스트
        st.markdown("**사업장 목록**")
        if not df_cur.empty and name_col:
            list_html = ""
            for _, row in df_cur.head(50).iterrows():
                nm  = str(row.get(name_col,""))
                typ = str(row.get(type_col,"")) if type_col else ""
                ind = str(row.get(ind_col,""))[:14] if ind_col else ""
                emp_v = int(row.get(emp_col,0)) if emp_col and pd.notna(row.get(emp_col)) else 0
                pill  = '<span class="pill-c">법인</span>' if "법인" in typ or typ=="1" else '<span class="pill-i">개인</span>'
                list_html += f"""
                <div class="list-item">
                  <div><div class="item-name">{nm}</div>
                  <div class="item-sub">{pill} {ind}</div></div>
                  <div class="item-emp">{emp_v:,}<span>명</span></div>
                </div>"""
            st.markdown(list_html, unsafe_allow_html=True)

        st.divider()

        # ── 핵심: 버튼 하나 ──────────────────────────────
        if gu != "선택" and not df_cur.empty:
            btn = st.button(f"⬇️ {gu} 주소·전번 보완 후 다운로드", type="primary")

            if btn:
                target = df_cur.copy()
                for col in ["전화번호","위도","경도"]:
                    if col not in target.columns:
                        target[col] = "" if col=="전화번호" else None
                    else:
                        if col=="전화번호": target[col] = target[col].fillna("").astype(str).replace("nan","")
                        else: target[col] = pd.to_numeric(target[col], errors="coerce")

                prog   = st.progress(0)
                status = st.empty()
                total2 = len(target)
                phone_ok = coord_ok = 0

                for i,(idx,row) in enumerate(target.iterrows()):
                    name   = str(row.get(name_col,"")) if name_col else ""
                    addr   = str(row.get(addr_col,"")) if addr_col else ""
                    biz_no = str(row.get(biz_col,"")) if biz_col else ""
                    is_corp= str(row.get(type_col,"")) in ["법인","1"] if type_col else False
                    phone=""; lat=None; lon=None

                    status.markdown(f"**[{i+1}/{total2}]** `{name[:20]}` — FSC→카카오→네이버")

                    if is_corp:
                        info = fsc(name, biz_no)
                        if info:
                            phone = info.get("phone","")
                            if info.get("addr") and addr_col: addr = info["addr"]
                        time.sleep(0.15)

                    kk = kakao_kw(name, addr)
                    if kk:
                        if not phone: phone = kk.get("phone","")
                        lat = float(kk.get("y") or 0) or None
                        lon = float(kk.get("x") or 0) or None
                        road = kk.get("road_address_name","")
                        if road and addr_col: addr = road

                    if not phone: phone = naver_phone(name, addr)

                    target.loc[idx,"전화번호"] = str(phone)
                    target.loc[idx,"위도"]    = lat
                    target.loc[idx,"경도"]    = lon
                    if addr_col: target.loc[idx,addr_col] = addr

                    if phone: phone_ok += 1
                    if lat:   coord_ok += 1
                    prog.progress((i+1)/total2)
                    time.sleep(random.uniform(0.2,0.4))

                # 컬럼 정리 + 불완전 행 삭제
                col_map = {}
                if name_col: col_map[name_col]="사업장명"
                if type_col: col_map[type_col]="법인여부"
                if addr_col: col_map[addr_col]="주소"
                col_map["전화번호"]="전화번호"
                if emp_col:  col_map[emp_col]="가입자수"
                if pay_col:  col_map[pay_col]="당월고지금액"
                col_map["위도"]="위도"; col_map["경도"]="경도"

                out_cols = [c for c in col_map if c in target.columns]
                out = target[out_cols].rename(columns=col_map)
                before = len(out)
                out = out[
                    out["전화번호"].notna() & (out["전화번호"]!="") &
                    out["주소"].notna()     & (out["주소"]!="")
                ].reset_index(drop=True)
                after = len(out)

                status.empty(); prog.empty()
                st.success(f"✅ 전화번호 **{phone_ok:,}개** | 좌표 **{coord_ok:,}개** | 최종 **{after:,}개**")

                fname = f"국민연금_{sido}_{gu}_보완완료.xlsx"
                st.download_button("⬇️ 엑셀 저장", make_excel(out), fname,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)

    else:
        st.info("왼쪽에서 국민연금 파일을 업로드하세요")

with col_main:
    # ── 탭: 지도 / 상세 ──────────────────────────────────
    tab_map, tab_list = st.tabs(["🗺️ 지도", "📋 상세"])

    with tab_map:
        if 'df_cur' in dir() and not df_cur.empty:
            df_map = df_cur[df_cur["위도"].notna() & df_cur["경도"].notna()] if "위도" in df_cur.columns else pd.DataFrame()
            lat_c  = float(df_map["위도"].mean()) if not df_map.empty else 36.5
            lon_c  = float(df_map["경도"].mean()) if not df_map.empty else 127.5
            m = folium.Map(location=[lat_c, lon_c], zoom_start=13 if not df_map.empty else 7,
                           tiles="CartoDB positron")
            if not df_map.empty:
                for _, r in df_map.iterrows():
                    size = max(6, min(28, int(r.get(emp_col,10)**0.5*2))) if emp_col else 8
                    color = "#3b82f6" if "법인" in str(r.get(type_col,"")) else "#f59e0b"
                    popup = folium.Popup(
                        f"<b>{r.get(name_col,'')}</b><br>📍 {str(r.get(addr_col,''))[:30]}<br>👥 {int(r.get(emp_col,0) if pd.notna(r.get(emp_col)) else 0)}명",
                        max_width=200)
                    folium.CircleMarker([r["위도"],r["경도"]], radius=size,
                        color=color, fill=True, fill_opacity=0.7, popup=popup).add_to(m)
            else:
                st.caption("⚠️ 좌표 데이터 없음 — 보완 후 다운로드 시 지도에 표시됩니다")
            st_folium(m, width="100%", height=750)
        else:
            st.markdown("""
            <div style="height:700px;display:flex;align-items:center;justify-content:center;flex-direction:column;color:#868e96;gap:10px">
              <div style="font-size:48px;opacity:.3">🗺️</div>
              <p style="font-size:14px;font-weight:600">파일을 업로드하고 지역을 선택하세요</p>
              <p style="font-size:12px">지도에 사업장이 표시됩니다</p>
            </div>""", unsafe_allow_html=True)

    with tab_list:
        if 'df_cur' in dir() and not df_cur.empty:
            show_cols = [c for c in [name_col,type_col,addr_col,emp_col,pay_col,ind_col] if c and c in df_cur.columns]
            st.dataframe(df_cur[show_cols].reset_index(drop=True), use_container_width=True, height=730)
        else:
            st.info("지역을 선택하면 상세 목록이 표시됩니다")
