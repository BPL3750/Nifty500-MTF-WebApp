import os
import io
import json
import pandas as pd
import streamlit as st
import yfinance as yf
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

st.set_page_config(page_title="Nifty 500 MTF & MIS Pro", page_icon="📈", layout="centered")

NIFTY_FILE = "Nifty_500_Stocks.xlsx"
MARGIN_FILE = "Zerodha_Dhan Margin Calculator.xlsx"
DEFAULT_DRIVE_FOLDER_ID = "11vyLILBunUtCb0CDWDx2yTLtYVOShatO"


@st.cache_data(ttl=3600)
def load_stocks_and_margins():
    stocks = []
    dhan_margins = {}
    zerodha_margins = {}

    if os.path.exists(MARGIN_FILE):
        xl_m = pd.ExcelFile(MARGIN_FILE)
        for s in xl_m.sheet_names:
            raw_df = pd.read_excel(xl_m, sheet_name=s, header=None, dtype=str)
            for idx, row in raw_df.iterrows():
                vals = [str(x).strip() if pd.notna(x) else "" for x in row.tolist()]
                if "Symbol / Scrip Name" in vals and "Margin Requirement User's Funds (in %)" in vals:
                    si = vals.index("Symbol / Scrip Name")
                    mi = vals.index("Margin Requirement User's Funds (in %)")
                    for _, r in raw_df.iloc[idx + 1:].iterrows():
                        rv = [str(x).strip() if pd.notna(x) else "" for x in r.tolist()]
                        if len(rv) > max(si, mi) and rv[si]:
                            try:
                                if rv[si] not in dhan_margins:
                                    dhan_margins[rv[si]] = float(rv[mi])
                            except ValueError:
                                pass
                elif "tradingsymbol" in vals and "margin" in vals and "isin" in vals:
                    si = vals.index("tradingsymbol")
                    mi = vals.index("margin")
                    for _, r in raw_df.iloc[idx + 1:].iterrows():
                        rv = [str(x).strip() if pd.notna(x) else "" for x in r.tolist()]
                        if len(rv) > max(si, mi) and rv[si]:
                            try:
                                zerodha_margins[rv[si]] = float(rv[mi])
                            except ValueError:
                                pass

    if os.path.exists(NIFTY_FILE):
        xl_n = pd.ExcelFile(NIFTY_FILE)
        for s in xl_n.sheet_names:
            raw_df = pd.read_excel(xl_n, sheet_name=s, header=None, dtype=str)
            for idx, row in raw_df.iterrows():
                vals = [str(x).strip() if pd.notna(x) else "" for x in row.tolist()]
                if "Company Name" in vals and "Symbol" in vals and "ISIN Code" in vals:
                    sub = raw_df.iloc[idx + 1:].copy()
                    sub.columns = vals
                    sub = sub[sub["Symbol"].notna() & (sub["Symbol"].str.strip() != "")]
                    for _, r in sub.iterrows():
                        sym = str(r["Symbol"]).strip()
                        ser = str(r.get("Series", "EQ")).strip()
                        stocks.append({
                            "Symbol": sym,
                            "Company": str(r["Company Name"]).strip(),
                            "Industry": str(r["Industry"]).strip(),
                            "Dhan_MTF": 100.0 if ser == "BE" else dhan_margins.get(sym, 30.0),
                            "Zerodha_MTF": 100.0 if ser == "BE" else zerodha_margins.get(sym, 30.0),
                            "MIS": 100.0 if ser == "BE" else 20.0,
                        })
                    break
            if stocks:
                break

    return pd.DataFrame(stocks)


@st.cache_data(ttl=300)
def get_live_cmp(symbol):
    try:
        hist = yf.Ticker(f"{symbol}.NS").history(period="5d")
        if not hist.empty:
            return round(float(hist["Close"].iloc[-1]), 2)
    except Exception:
        pass
    return 0.0


@st.cache_data(ttl=900)
def fetch_latest_gdrive_reports():
    """Streamlit Secrets मधून सुरक्षितपणे Google Drive वरील ताजे स्कॅन रिपोर्ट्स वाचणे"""
    if "GDRIVE_TOKEN_JSON" not in st.secrets:
        return None, [], "Google Drive जोडलेले नाही (खालील सूचनांनुसार Streamlit Secrets मध्ये Token टाका)."

    try:
        token_info = json.loads(st.secrets["GDRIVE_TOKEN_JSON"])
        root_id = st.secrets.get("GDRIVE_FOLDER_ID", DEFAULT_DRIVE_FOLDER_ID)
        scopes = ["https://googleapis.com"]
        creds = Credentials.from_authorized_user_info(token_info, scopes)
        service = build("drive", "v3", credentials=creds)

        # 01_Daily_Reports फोल्डर शोधणे
        q_daily = f"name='01_Daily_Reports' and '{root_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        daily_res = service.files().list(q=q_daily, fields="files(id)").execute().get("files", [])
        if not daily_res:
            return None, [], "Drive वर '01_Daily_Reports' फोल्डर सापडले नाही."
        daily_id = next(iter(daily_res))["id"]

        # सर्वात ताजे तारीख फोल्डर शोधणे
        q_dates = f"'{daily_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        date_folders = service.files().list(
            q=q_dates, orderBy="createdTime desc", pageSize=5, fields="files(id, name)"
        ).execute().get("files", [])
        if not date_folders:
            return None, [], "अजून कोणतेही तारीख फोल्डर तयार झालेले नाही."

        latest_folder = next(iter(date_folders))

        folder_ids = [latest_folder["id"]]

        # आतील सब-फोल्डर्स (Full_Market_Scan, Sector_and_Breadth इ.) शोधणे
        q_subs = f"'{latest_folder['id']}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        sub_folders = service.files().list(q=q_subs, fields="files(id, name)").execute().get("files", [])
        for sf in sub_folders:
            folder_ids.append(sf["id"])

        report_files = []
        for fid in folder_ids:
            q_files = f"'{fid}' in parents and trashed=false and (name contains '.xlsx' or name contains '.csv')"
            items = service.files().list(q=q_files, fields="files(id, name)").execute().get("files", [])
            for it in items:
                if "Calculator" not in it["name"]:
                    report_files.append(it)

        return latest_folder["name"], report_files, service
    except Exception as e:
        return None, [], f"Drive कनेक्शन त्रुटी: {e}"

def calculate_trade(broker, product, capital, entry, sl, t1, t2, days, risk_pct, margin_pct):
    m_dec = margin_pct / 100.0
    leverage = round(1.0 / m_dec, 2) if m_dec > 0 else 1.0
    buying_power = capital / m_dec if m_dec > 0 else 0.0
    max_qty = int(buying_power // entry) if entry > 0 else 0

    sl_diff = abs(entry - sl)
    risk_amt = capital * (risk_pct / 100.0)
    safe_qty = int(risk_amt // sl_diff) if sl_diff > 0 else 0

    trade_val = safe_qty * entry
    margin_used = trade_val * m_dec
    funded_amt = max(0.0, trade_val - margin_used) if product == "MTF Swing" else 0.0
    danger_price = round(entry * (1 - m_dec * 0.5), 2) if product == "MTF Swing" else round(entry * (1 - m_dec * 0.6), 2)

    sell_val = safe_qty * t2
    if safe_qty == 0:
        return {
            "leverage": leverage, "buying_power": buying_power, "max_qty": max_qty,
            "safe_qty": 0, "trade_val": 0.0, "margin_used": 0.0, "funded_amt": 0.0,
            "danger_price": danger_price, "brokerage": 0.0, "interest": 0.0,
            "taxes": 0.0, "total_charges": 0.0, "breakeven": entry,
            "net_profit": 0.0, "net_loss": 0.0, "roi": 0.0,
            "half_qty": 0, "rem_qty": 0, "p1_profit": 0.0, "p2_profit": 0.0, "partial_total": 0.0
        }

    if product == "MTF Swing":
        stt = round(trade_val * 0.001) + round(sell_val * 0.001)
        stamp = round(trade_val * 0.00015)
        dp_pledge = (14.75 if broker == "Dhan" else 15.34) + 35.40
        rate = 0.000342 if broker == "Dhan" else 0.0004
        interest = round(funded_amt * rate * days, 2)
        if broker == "Zerodha":
            brokerage = min(20.0, trade_val * 0.003) + min(20.0, sell_val * 0.003)
        else:
            brokerage = min(20.0, trade_val * 0.0003) + min(20.0, sell_val * 0.0003)
    else:
        stt = round(sell_val * 0.00025)
        stamp = round(trade_val * 0.00003)
        dp_pledge = 0.0
        interest = 0.0
        brokerage = min(20.0, trade_val * 0.0003) + min(20.0, sell_val * 0.0003)

    exch_sebi = (trade_val + sell_val) * 0.0000317
    gst = round((brokerage + exch_sebi) * 0.18, 2)
    taxes = round(stt + stamp + exch_sebi + gst + dp_pledge, 2)
    total_charges = round(brokerage + interest + taxes, 2)

    breakeven = round(entry + (total_charges / safe_qty), 2)
    net_profit = round(((t2 - entry) * safe_qty) - total_charges, 2)
    net_loss = round(((sl - entry) * safe_qty) - total_charges, 2)
    roi = round((net_profit / margin_used) * 100, 2) if margin_used > 0 else 0.0

    # ५०% अर्धा नफा बुकिंग गणना
    half_qty = safe_qty // 2
    rem_qty = safe_qty - half_qty
    half_cost = round(total_charges * 0.5, 2)
    p1_profit = round(((t1 - entry) * half_qty) - half_cost, 2) if half_qty > 0 else 0.0
    p2_profit = round(((t2 - entry) * rem_qty) - half_cost, 2) if rem_qty > 0 else 0.0
    partial_total = round(p1_profit + p2_profit, 2)

    return {
        "leverage": leverage, "buying_power": buying_power, "max_qty": max_qty,
        "safe_qty": safe_qty, "trade_val": trade_val, "margin_used": margin_used,
        "funded_amt": funded_amt, "danger_price": danger_price,
        "brokerage": round(brokerage, 2), "interest": interest,
        "taxes": taxes, "total_charges": total_charges, "breakeven": breakeven,
        "net_profit": net_profit, "net_loss": net_loss, "roi": roi,
        "half_qty": half_qty, "rem_qty": rem_qty,
        "p1_profit": p1_profit, "p2_profit": p2_profit, "partial_total": partial_total
    }

def render_app():
    st.title("📈 Nifty 500 MTF & MIS Pro")
    st.caption("Dhan आणि Zerodha पोझिशन सायझिंग, ५०% प्रॉफिट बुकिंग, Trailing SL व लाइव्ह ब्रेकआऊट स्कॅनर")

    df_stocks = load_stocks_and_margins()
    if df_stocks.empty:
        st.error("Nifty_500_Stocks.xlsx सापडली नाही! कृपया GitHub रेपॉझिटरी तपासा.")
        return

    tab1, tab2, tab3 = st.tabs([
        "🧮 MTF / MIS कॅल्क्युलेटर",
        "🚀 आजचे ब्रेकआऊट रिपोर्ट्स",
        "📊 ५०० शेअर्स मार्जिन यादी"
    ])

    with tab1:
        col_b, col_p = st.columns(2)
        broker = col_b.selectbox("ब्रोकर निवडा", ["Dhan", "Zerodha"])
        product = col_p.selectbox("प्रॉडक्ट प्रकार", ["MTF Swing", "MIS Intraday"])

        symbols = df_stocks["Symbol"].tolist()
        default_idx = symbols.index("RELIANCE") if "RELIANCE" in symbols else 0
        selected_sym = st.selectbox("🔍 Nifty 500 शेअर निवडा (टाईप करून शोधा)", symbols, index=default_idx)

        stock_row = df_stocks[df_stocks["Symbol"] == selected_sym].iloc[0]
        live_cmp = get_live_cmp(selected_sym)

        st.info(f"**{stock_row['Company']}** ({stock_row['Industry']}) | लाइव्ह भाव: **₹{live_cmp:,.2f}**")

        c1, c2 = st.columns(2)
        capital = c1.number_input("तुमचे भांडवल (₹)", min_value=1000.0, value=10000.0, step=1000.0)
        risk_pct = c2.number_input("प्रति ट्रेड जोखीम (%)", min_value=0.1, max_value=20.0, value=2.0, step=0.5)

        default_entry = live_cmp if live_cmp > 0 else 1300.0
        c3, c4 = st.columns(2)
        entry = c3.number_input("खरेदी किंमत Entry (₹)", min_value=1.0, value=float(default_entry), step=1.0)
        sl = c4.number_input("स्टॉप लॉस SL (₹)", min_value=0.5, value=round(float(default_entry) * 0.97, 2), step=1.0)

        c5, c6, c7 = st.columns(3)
        t1 = c5.number_input("Target 1 (50% Book)", min_value=1.0, value=round(float(default_entry) * 1.05, 2), step=1.0)
        t2 = c6.number_input("Target 2 (Final)", min_value=1.0, value=round(float(default_entry) * 1.10, 2), step=1.0)
        days = c7.number_input("होल्डिंग दिवस", min_value=0, max_value=365, value=10 if product == "MTF Swing" else 0)

        dhan_m = stock_row["MIS"] if product == "MIS Intraday" else stock_row["Dhan_MTF"]
        zero_m = stock_row["MIS"] if product == "MIS Intraday" else stock_row["Zerodha_MTF"]

        res_dhan = calculate_trade("Dhan", product, capital, entry, sl, t1, t2, days, risk_pct, dhan_m)
        res_zero = calculate_trade("Zerodha", product, capital, entry, sl, t1, t2, days, risk_pct, zero_m)
        active = res_dhan if broker == "Dhan" else res_zero

        st.divider()
        st.subheader(f"🎯 {broker} ({product}) निकाल")

        m1, m2, m3 = st.columns(3)
        m1.metric("Safe Qty (२% रिस्क)", f"{active['safe_qty']} शेअर्स", f"Max Qty: {active['max_qty']}")
        m2.metric("लीव्हरेज / मार्जिन", f"{active['leverage']}x", f"{dhan_m if broker=='Dhan' else zero_m}%")
        m3.metric("तुमचे वापरलेले भांडवल", f"₹{active['margin_used']:,.0f}", f"खरेदी: ₹{active['trade_val']:,.0f}")

        m4, m5, m6 = st.columns(3)
        m4.metric("Target 2 निव्वळ नफा", f"₹{active['net_profit']:,.2f}", f"ROI: {active['roi']}%")
        m5.metric("SL वर प्रत्यक्ष तोटा", f"₹{active['net_loss']:,.2f}", f"Breakeven: ₹{active['breakeven']}")
        m6.metric("एकूण कर व व्याज", f"₹{active['total_charges']:,.2f}", f"व्याज: ₹{active['interest']}")

        if product == "MTF Swing":
            st.warning(f"🚨 **MTF Margin Call धोक्याची किंमत:** ₹{active['danger_price']:,.2f} (फंडिंग रक्कम: ₹{active['funded_amt']:,.0f})")

        st.subheader("💰 ५०% अर्धा नफा बुकिंग प्लॅनर (Partial Profit Booking)")
        partial_df = pd.DataFrame({
            "टप्पा (Stage)": ["टप्पा १: Target 1 (50% Book)", "टप्पा २: Target 2 (उरलेले 50%)", "एकूण एकत्रित निव्वळ नफा (Total)"],
            "विक्री किंमत (₹)": [f"₹{t1:,.2f}", f"₹{t2:,.2f}", "-"],
            "शेअर संख्या (Qty)": [f"{active['half_qty']} शेअर्स", f"{active['rem_qty']} शेअर्स", f"{active['safe_qty']} शेअर्स"],
            "निव्वळ नफा (Net Profit ₹)": [f"₹{active['p1_profit']:,.2f}", f"₹{active['p2_profit']:,.2f}", f"₹{active['partial_total']:,.2f}"]
        })
        st.dataframe(partial_df, use_container_width=True, hide_index=True)

        st.subheader("⚔️ Dhan विरुद्ध Zerodha थेट तुलना (Safe Qty)")
        comp_df = pd.DataFrame({
            "तपशील": ["मार्जिन %", "ब्रोकरेज (₹)", "MTF व्याज (₹)", "सरकारी कर व DP (₹)", "एकूण खर्च (₹)", "Target 2 निव्वळ नफा (₹)"],
            "Dhan": [f"{dhan_m}%", res_dhan["brokerage"], res_dhan["interest"], res_dhan["taxes"], res_dhan["total_charges"], res_dhan["net_profit"]],
            "Zerodha": [f"{zero_m}%", res_zero["brokerage"], res_zero["interest"], res_zero["taxes"], res_zero["total_charges"], res_zero["net_profit"]],
        })
        st.dataframe(comp_df, use_container_width=True, hide_index=True)

        diff = round(res_dhan["net_profit"] - res_zero["net_profit"], 2)
        if diff > 0:
            st.success(f"✅ **Dhan** वर तुम्हाला **₹{diff:,.2f} जास्त निव्वळ नफा** उरतो!")
        elif diff < 0:
            st.success(f"✅ **Zerodha** वर तुम्हाला **₹{abs(diff):,.2f} जास्त निव्वळ नफा** उरतो!")

        st.subheader("🛡️ Trailing Stop-Loss (TSL) प्लॅनर")
        risk_pts = abs(entry - sl)
        tsl_df = pd.DataFrame({
            "टप्पा": ["1R (1:1)", "1.5R (1:1.5)", "2R (1:2)", "3R (1:3)"],
            "शेअर किंमत (Trigger ₹)": [round(entry + risk_pts, 2), round(entry + 1.5 * risk_pts, 2), round(entry + 2 * risk_pts, 2), round(entry + 3 * risk_pts, 2)],
            "नवीन Trailing SL (₹)": [active["breakeven"], round(entry + 0.5 * risk_pts, 2), round(entry + risk_pts, 2), round(entry + 2 * risk_pts, 2)],
        })
        st.dataframe(tsl_df, use_container_width=True, hide_index=True)

    with tab2:
        st.subheader("🚀 Google Drive वरील आजचे ब्रेकआऊट आणि सेक्टर रिपोर्ट्स")
        folder_date, rep_files, service_or_msg = fetch_latest_gdrive_reports()

        if not folder_date:
            st.info(service_or_msg)
            st.markdown("""
            **Google Drive जोडण्यासाठी (फक्त १ मिनिटाचे काम):**
            1. तुमची रेपॉझिटरी **Public** असल्यामुळे कधीही `token.json` फाईल GitHub वर टाकू नका.
            2. त्याऐवजी **Streamlit Cloud -> तुमच्या App चे `Settings` -> `Secrets`** उघडा.
            3. तेथे तुमच्या `C:\\NSE-Stock-Automation\\token.json` मधील मजकूर खालीलप्रमाणे पेस्ट करा:
            ```toml
            GDRIVE_FOLDER_ID = "11vyLILBunUtCb0CDWDx2yTLtYVOShatO"
            GDRIVE_TOKEN_JSON = '''येथे_तुमचा_token.json_मजकूर_पेस्ट_करा'''
            ```
            """)
        elif not rep_files:
            st.warning(f"📅 **{folder_date}** फोल्डरमध्ये कोणताही Excel/CSV रिपोर्ट सापडला नाही.")
        else:
            st.success(f"📅 ताजे रिपोर्ट्स लोड झाले आहेत: **{folder_date}**")
            file_map = {f["name"]: f["id"] for f in rep_files}
            chosen_name = st.selectbox("रिपोर्ट निवडा", list(file_map.keys()))

            if chosen_name:
                request = service_or_msg.files().get_media(fileId=file_map[chosen_name])
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                fh.seek(0)

                if chosen_name.endswith(".csv"):
                    rep_df = pd.read_csv(fh)
                    st.dataframe(rep_df, use_container_width=True, hide_index=True)
                else:
                    xl_rep = pd.ExcelFile(fh)
                    sheet_sel = st.selectbox("शीट (Tab) निवडा", xl_rep.sheet_names)
                    rep_df = pd.read_excel(xl_rep, sheet_name=sheet_sel)
                    st.dataframe(rep_df, use_container_width=True, hide_index=True)

    with tab3:
        st.subheader("सर्व Nifty 500 शेअर्सची Dhan व Zerodha MTF यादी")
        search_q = st.text_input("कंपनी किंवा सेक्टर शोधा (उदा. BANK, TATA, PHARMA)")
        view_df = df_stocks.copy()
        if search_q:
            view_df = view_df[
                view_df["Symbol"].str.contains(search_q, case=False) |
                view_df["Company"].str.contains(search_q, case=False) |
                view_df["Industry"].str.contains(search_q, case=False)
            ]
        st.dataframe(view_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    render_app()
