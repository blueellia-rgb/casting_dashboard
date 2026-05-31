"""
국민연금 사업장 주소·전번 보완기
파일 업로드 → 자동 보완 → 엑셀 다운로드
"""
import io, re, time, random, requests
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="사업장 주소·전번 보완",
    page_icon="📋",
    layout="centered"
)

st.markdown("""
<style>
#MainMenu,footer,header{visibility:hidden}
.block-container{max-width:680px!important;padding:2rem 1rem!important}
.step-box{background:#f8f9fa;border-radius:12px;padding:20px;margin-bottom:16px;border:1px solid #dee2e6}
.step-num{background:#3b82f6;color:#fff;border-radius:50%;width:24px;height:24px;display:inline-flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;margin-right:8px}
.result-box{background:#ecfdf5;border-radius:12px;padding:16px;border:1px solid #a7f3d0;margin-top:16px}
div.stButton>button{background:#3b82f6!important;color:#fff!important;border:none!important;border-radius:8px!important;font-weight:600!important;font-size:14px!important;padding:10px 0!important;width:100%!important}
div.stDownloadButton>button{background:#10b981!important;color:#fff!important;border:none!important;border-radius:8px!important;font-weight:600!important;font-size:14px!important;padding:10px 0!important;width:100%!important}
</style>
""", unsafe_allow_html=True)

# ── API 키 ─────────────────────────────────────────────
FSC_KEY   = "2446da601452d231f1ff77a4c5ac9598896028fb96e8be586bf7f4be58b6231c"
KAKAO_KEY = "e8eef7246a2cd0f3d2d55af8a806e3ef"
NAV_HDR   = {
    "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15",
    "Referer": "https://map.naver.com/",
    "Accept-Language": "ko-KR,ko;q=0.9"
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
    gu = re.search(r'[\w]+구', addr)
    gu = gu.group() if gu else ""
    try:
        r = requests.get("https://m.map.naver.com/search2/search.nhn",
            headers=NAV_HDR, params={"query":f"{clean} {gu}","type":"all"}, timeout=8)
        if r.status_code == 200:
            for pat in [r'tel:(0\d{1,2}-?\d{3,4}-?\d{4})',
                        r'(?<!\d)(0\d{1,2}-\d{3,4}-\d{4})(?!\d)']:
                f = re.findall(pat, r.text)
                if f: return f[0]
    except: pass
    return ""

def make_excel(df):
    wb = Workbook(); ws = wb.active; ws.title = "보완완료"
    def thin():
        s = Side(style="thin", color="D0D0D0")
        return Border(left=s,right=s,top=s,bottom=s)
    hdrs = list(df.columns)
    widths = {"사업장명":26,"업종":20,"법인여부":8,"주소":44,"전화번호":14,
              "가입자수":8,"당월고지금액":14,"위도":12,"경도":12}
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

# ══════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════
st.markdown("# 📋 사업장 주소·전번 보완")
st.caption("국민연금 대시보드에서 다운로드한 파일을 업로드하면 전화번호와 상세주소를 자동으로 추가합니다.")
st.divider()

# STEP 1
st.markdown('<div class="step-box"><span class="step-num">1</span><b>파일 업로드</b><br><small style="color:#868e96">HTML 대시보드에서 다운로드한 엑셀 파일을 올려주세요</small></div>', unsafe_allow_html=True)
uploaded = st.file_uploader("파일 선택", type=["xlsx"], label_visibility="collapsed")

if uploaded:
    df = pd.read_excel(uploaded)
    df.columns = [c.strip() for c in df.columns]

    # 컬럼 감지
    name_col = next((c for c in df.columns if "사업장명" in c or "상호" in c), None)
    addr_col = next((c for c in df.columns if "주소" in c or "도로명" in c), None)
    biz_col  = next((c for c in df.columns if "사업자" in c and "번호" in c), None)
    type_col = next((c for c in df.columns if "형태" in c or "법인" in c), None)
    emp_col  = next((c for c in df.columns if "가입자" in c), None)
    pay_col  = next((c for c in df.columns if "고지" in c), None)
    ind_col  = next((c for c in df.columns if "업종" in c), None)

    # 컬럼 초기화
    for col in ["전화번호"]:
        if col not in df.columns: df[col] = ""
        else: df[col] = df[col].fillna("").astype(str).replace("nan","")
    for col in ["위도","경도"]:
        if col not in df.columns: df[col] = None
        else: df[col] = pd.to_numeric(df[col], errors="coerce")

    total = len(df)
    region = uploaded.name.replace(".xlsx","")

    # 파일 정보
    corp = int((df[type_col].astype(str).str.contains("법인|^1$")).sum()) if type_col else 0
    st.success(f"✅ **{region}** — 총 {total:,}개 사업장 (법인 {corp:,}개)")
    st.dataframe(df.head(5)[[c for c in [name_col,addr_col,emp_col] if c]], use_container_width=True, height=200)

    st.divider()

    # STEP 2
    st.markdown('<div class="step-box"><span class="step-num">2</span><b>자동 보완 시작</b><br><small style="color:#868e96">FSC 금융위원회 → 카카오 → 네이버 순으로 전화번호·주소를 수집합니다</small></div>', unsafe_allow_html=True)

    est = int(total * 0.4 / 60)
    st.caption(f"⏱ 예상 소요시간: 약 {est}분 ({total:,}개 × 0.4초)")

    if st.button("🚀 주소·전번 보완 시작", type="primary"):
        prog   = st.progress(0, text="준비 중...")
        status = st.empty()
        log    = st.empty()

        phone_ok = coord_ok = 0

        for i,(idx,row) in enumerate(df.iterrows()):
            name   = str(row.get(name_col,"")) if name_col else ""
            addr   = str(row.get(addr_col,"")) if addr_col else ""
            biz_no = str(row.get(biz_col,"")) if biz_col else ""
            is_corp= str(row.get(type_col,"")) in ["법인","1"] if type_col else False
            phone=""; lat=None; lon=None

            pct = (i+1)/total
            prog.progress(pct, text=f"{i+1:,} / {total:,}개 처리 중 ({pct*100:.0f}%)")
            status.markdown(f"🔍 `{name[:25]}`")

            # ① FSC (법인)
            if is_corp:
                info = fsc(name, biz_no)
                if info:
                    phone = info.get("phone","")
                    if info.get("addr") and addr_col: addr = info["addr"]
                time.sleep(0.15)

            # ② 카카오
            kk = kakao_kw(name, addr)
            if kk:
                if not phone: phone = kk.get("phone","")
                lat = float(kk.get("y") or 0) or None
                lon = float(kk.get("x") or 0) or None
                road = kk.get("road_address_name","")
                if road and addr_col: addr = road

            # ③ 네이버
            if not phone:
                phone = naver_phone(name, addr)

            df.loc[idx,"전화번호"] = str(phone)
            df.loc[idx,"위도"]    = lat
            df.loc[idx,"경도"]    = lon
            if addr_col: df.loc[idx,addr_col] = addr

            if phone: phone_ok += 1
            if lat:   coord_ok += 1

            # 10개마다 중간 현황
            if (i+1) % 10 == 0:
                log.markdown(f"📞 전화번호 **{phone_ok}개** | 📍 좌표 **{coord_ok}개** 수집됨")

            time.sleep(random.uniform(0.2, 0.4))

        # 완료
        prog.empty(); status.empty(); log.empty()

        # 컬럼 정리 + 불완전 행 삭제
        col_map = {}
        if name_col: col_map[name_col]="사업장명"
        if type_col: col_map[type_col]="법인여부"
        if addr_col: col_map[addr_col]="주소"
        col_map["전화번호"]="전화번호"
        if emp_col:  col_map[emp_col]="가입자수"
        if pay_col:  col_map[pay_col]="당월고지금액"
        if ind_col:  col_map[ind_col]="업종"
        col_map["위도"]="위도"; col_map["경도"]="경도"

        out_cols = [c for c in col_map if c in df.columns]
        out = df[out_cols].rename(columns=col_map)
        before = len(out)
        out = out[
            out["전화번호"].notna() & (out["전화번호"]!="") &
            out["주소"].notna()     & (out["주소"]!="")
        ].reset_index(drop=True)
        after = len(out)

        st.balloons()
        st.markdown(f"""
        <div class="result-box">
          ✅ <b>보완 완료!</b><br><br>
          📞 전화번호 수집: <b>{phone_ok:,}개</b> / {total:,}개<br>
          📍 좌표 수집: <b>{coord_ok:,}개</b> / {total:,}개<br>
          🗑️ 삭제된 행: <b>{before-after:,}개</b> (전화번호·주소 없음)<br>
          📊 최종 저장: <b>{after:,}개</b>
        </div>
        """, unsafe_allow_html=True)

        fname = region + "_보완완료.xlsx"
        st.download_button(
            "⬇️ 보완 완료 엑셀 다운로드",
            make_excel(out), fname,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
