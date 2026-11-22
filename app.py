import os
import io
import json
import pandas as pd
import streamlit as st
import plotly.express as px
import yfinance as yf
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

st.set_page_config(page_title="Nifty 500 MTF & MIS Pro", page_icon="📈", layout="wide")

NIFTY_FILE = "Nifty_500_Stocks.xlsx"
MARGIN_FILE = "Zerodha_Dhan Margin Calculator.xlsx"
DEFAULT_DRIVE_FOLDER_ID = "11vyLILBunUtCb0CDWDx2yTLtYVOShatO"


def get_gdrive_service():
    if "GDRIVE_TOKEN_JSON" not in st.secrets:
        return None, None, "Google Drive जोडलेले नाही (Streamlit Secrets मध्ये Token तपासा)."
    try:
        token_info = json.loads(st.secrets["GDRIVE_TOKEN_JSON"])
        root_id = st.secrets.get("GDRIVE_FOLDER_ID", DEFAULT_DRIVE_FOLDER_ID)
        drive_scope = "".join(["https://", "www.", "googleapis.", "com/", "auth/", "drive"])
        creds = Credentials.from_authorized_user_info(token_info, [drive_scope])
        service = build("drive", "v3", credentials=creds)
        return service, root_id, None
    except Exception as e:
        return None, None, f"Drive कनेक्शन त्रुटी: {e}"


@st.cache_data(ttl=900)
def fetch_all_live_prices(symbols_tuple):
    prices = {}
    try:
        tickers = [f"{s}.NS" for s in symbols_tuple if s and not s.startswith("DUMMY")]
        data = yf.download(tickers, period="2d", progress=False)
        if "Close" in data:
            close_df = data["Close"]
            for s in symbols_tuple:
                col_name = f"{s}.NS"
                if col_name in close_df.columns:
                    vals = close_df[col_name].dropna()
                    if not vals.empty:
                        prices[s] = round(float(vals.iloc[-1]), 2)
    except Exception:
        pass
    return prices

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

    df_res = pd.DataFrame(stocks)
    if not df_res.empty:
        live_map = fetch_all_live_prices(tuple(df_res["Symbol"].tolist()))
        df_res["Live_CMP"] = df_res["Symbol"].map(live_map).fillna(0.0)
    return df_res

@st.cache_data(ttl=900)
def fetch_latest_gdrive_reports():
    service, root_id, err = get_gdrive_service()
    if err:
        return None, [], err

    try:
        q_daily = f"name='01_Daily_Reports' and '{root_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        daily_res = service.files().list(q=q_daily, fields="files(id)").execute().get("files", [])
        if not daily_res:
            return None, [], "Drive वर '01_Daily_Reports' फोल्डर सापडले नाही."

        daily_id = next(iter(daily_res)).get("id")
        q_dates = f"'{daily_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        date_folders = service.files().list(
            q=q_dates, orderBy="createdTime desc", pageSize=5, fields="files(id, name)"
        ).execute().get("files", [])
        if not date_folders:
            return None, [], "अजून कोणतेही तारीख फोल्डर तयार झालेले नाही."

        latest_folder = next(iter(date_folders))
        folder_ids = [latest_folder.get("id")]

        q_subs = f"'{latest_folder.get('id')}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        sub_folders = service.files().list(q=q_subs, fields="files(id, name)").execute().get("files", [])
        for sf in sub_folders:
            folder_ids.append(sf.get("id"))

        report_files = []
        for fid in folder_ids:
            q_files = f"'{fid}' in parents and trashed=false and (name contains '.xlsx' or name contains '.csv')"
            items = service.files().list(q=q_files, fields="files(id, name)").execute().get("files", [])
            for it in items:
                if "Calculator" not in it.get("name", ""):
                    report_files.append(it)

        return latest_folder.get("name"), report_files, service
    except Exception as e:
        return None, [], f"Drive त्रुटी: {e}"


@st.cache_data(ttl=600)
def fetch_master_trade_journal():
    service, root_id, err = get_gdrive_service()
    if err:
        return None, err

    try:
        q_j = f"name='04_Trade_Journal' and '{root_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        j_res = service.files().list(q=q_j, fields="files(id)").execute().get("files", [])
        if not j_res:
            return None, "Drive वर '04_Trade_Journal' फोल्डर सापडले नाही."

        j_id = next(iter(j_res)).get("id")
        q_f = f"name='Master_Trade_Journal.xlsx' and '{j_id}' in parents and trashed=false"
        f_res = service.files().list(q=q_f, fields="files(id)").execute().get("files", [])
        if not f_res:
            return None, "अजून 'Master_Trade_Journal.xlsx' तयार झालेली नाही."

        file_id = next(iter(f_res)).get("id")
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)

        xl = pd.ExcelFile(fh)
        sheets_data = {s: pd.read_excel(xl, sheet_name=s) for s in xl.sheet_names}
        return sheets_data, None
    except Exception as e:
        return None, f"जर्नल वाचताना त्रुटी: {e}"

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

def render_tab1_calculator(df_stocks):
    col_b, col_p = st.columns(2)
    broker = col_b.selectbox("ब्रोकर निवडा", ["Dhan", "Zerodha"])
    product = col_p.selectbox("प्रॉडक्ट प्रकार", ["MTF Swing", "MIS Intraday"])

    symbols = df_stocks["Symbol"].tolist()
    default_idx = symbols.index("RELIANCE") if "RELIANCE" in symbols else 0
    selected_sym = st.selectbox("🔍 Nifty 500 शेअर निवडा (टाईप करून शोधा)", symbols, index=default_idx)

    stock_row = df_stocks[df_stocks["Symbol"] == selected_sym].to_dict("records").pop(0)
    live_cmp = float(stock_row.get("Live_CMP", 0.0))

    st.info(f"**{stock_row['Company']}** ({stock_row['Industry']}) | लाइव्ह भाव: **₹{live_cmp:,.2f}**")

    c1, c2 = st.columns(2)
    capital = c1.number_input("तुमचे भांडवल (₹)", min_value=1000.0, value=10000.0, step=1000.0)
    risk_pct = c2.number_input("प्रति ट्रेड जोखीम (%)", min_value=0.1, max_value=20.0, value=2.0, step=0.5)

    default_entry = live_cmp if live_cmp > 0 else 1300.0
    c3, c4 = st.columns(2)
    entry = c3.number_input("खरेदी किंमत Entry (₹)", min_value=1.0, value=float(default_entry), step=1.0)
    sl = c4.number_input("स्टॉप लॉस SL (₹)", min_value=0.5, value=round(float(default_entry) * 0.97, 2), step=1.0)

    c5, c6, c7 = st.columns(3)
    t1 = c5.number_input("Target 1 (50% Book ₹)", min_value=1.0, value=round(float(default_entry) * 1.05, 2), step=1.0)
    t2 = c6.number_input("Target 2 (Final ₹)", min_value=1.0, value=round(float(default_entry) * 1.10, 2), step=1.0)
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
    m3.metric("तुमचे वापरलेले भांडवल", f"₹{active['margin_used']:,.2f}", f"खरेदी: ₹{active['trade_val']:,.2f}")

    m4, m5, m6 = st.columns(3)
    m4.metric("Target 2 निव्वळ नफा", f"₹{active['net_profit']:,.2f}", f"ROI: {active['roi']}%")
    m5.metric("SL वर प्रत्यक्ष तोटा", f"₹{active['net_loss']:,.2f}", f"Breakeven: ₹{active['breakeven']:,.2f}")
    m6.metric("एकूण कर व व्याज", f"₹{active['total_charges']:,.2f}", f"व्याज: ₹{active['interest']:,.2f}")

    if product == "MTF Swing":
        st.warning(f"🚨 **MTF Margin Call धोक्याची किंमत:** ₹{active['danger_price']:,.2f} (फंडिंग रक्कम: ₹{active['funded_amt']:,.2f})")

    st.subheader("💰 ५०% अर्धा नफा बुकिंग प्लॅनर")
    partial_df = pd.DataFrame({
        "टप्पा (Stage)": ["टप्पा १: Target 1 (50% Book)", "टप्पा २: Target 2 (उरलेले 50%)", "एकूण एकत्रित निव्वळ नफा"],
        "विक्री किंमत": [f"₹{t1:,.2f}", f"₹{t2:,.2f}", "-"],
        "शेअर्स (Qty)": [f"{active['half_qty']} शेअर्स", f"{active['rem_qty']} शेअर्स", f"{active['safe_qty']} शेअर्स"],
        "निव्वळ नफा": [f"₹{active['p1_profit']:,.2f}", f"₹{active['p2_profit']:,.2f}", f"₹{active['partial_total']:,.2f}"]
    })
    st.dataframe(partial_df, use_container_width=True, hide_index=True)

    st.subheader("⚔️ Dhan विरुद्ध Zerodha थेट तुलना")
    comp_df = pd.DataFrame({
        "तपशील": ["मार्जिन %", "ब्रोकरेज", "MTF व्याज", "सरकारी कर व DP", "एकूण खर्च", "Target 2 निव्वळ नफा"],
        "Dhan": [f"{dhan_m}%", f"₹{res_dhan['brokerage']:,.2f}", f"₹{res_dhan['interest']:,.2f}", f"₹{res_dhan['taxes']:,.2f}", f"₹{res_dhan['total_charges']:,.2f}", f"₹{res_dhan['net_profit']:,.2f}"],
        "Zerodha": [f"{zero_m}%", f"₹{res_zero['brokerage']:,.2f}", f"₹{res_zero['interest']:,.2f}", f"₹{res_zero['taxes']:,.2f}", f"₹{res_zero['total_charges']:,.2f}", f"₹{res_zero['net_profit']:,.2f}"],
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
        "Trigger किंमत": [f"₹{entry + risk_pts:,.2f}", f"₹{entry + 1.5 * risk_pts:,.2f}", f"₹{entry + 2 * risk_pts:,.2f}", f"₹{entry + 3 * risk_pts:,.2f}"],
        "नवीन Trailing SL": [f"₹{active['breakeven']:,.2f}", f"₹{entry + 0.5 * risk_pts:,.2f}", f"₹{entry + risk_pts:,.2f}", f"₹{entry + 2 * risk_pts:,.2f}"],
    })
    st.dataframe(tsl_df, use_container_width=True, hide_index=True)

def get_smart_col_config(df):
    """रिपोर्टमधील किंमत आणि रकमेच्या कॉलम्सना आपोआप ₹ फॉरमॅट लावणे"""
    cfg = {}
    price_keywords = (
        "PRICE", "CMP", "TARGET", "SL", "ATH", "DMA", "HIGH", "LOW",
        "AMT", "AMOUNT", "VALUE", "MOVE", "GAIN", "LOSS", "PNL"
    )
    exclude_keywords = ("QTY", "SCORE", "RATIO", "DATE", "DAYS", "COUNT", "ID")

    for col in df.columns:
        u_col = str(col).strip().upper()
        if any(ex in u_col for ex in exclude_keywords):
            continue
        if "%" in u_col or "PCT" in u_col:
            cfg[col] = st.column_config.NumberColumn(format="%.2f%%")
        elif any(pk in u_col for pk in price_keywords):
            cfg[col] = st.column_config.NumberColumn(format="₹%.2f")
    return cfg


def render_tab2_and_tab3(tab2, tab3, df_stocks):
    with tab2:
        st.subheader("🚀 Google Drive वरील आजचे ब्रेकआऊट आणि सेक्टर रिपोर्ट्स")
        folder_date, rep_files, service_or_msg = fetch_latest_gdrive_reports()

        if not folder_date:
            st.info(service_or_msg)
        elif not rep_files:
            st.warning(f"📅 **{folder_date}** फोल्डरमध्ये कोणताही रिपोर्ट सापडला नाही.")
        else:
            st.success(f"📅 ताजे रिपोर्ट्स लोड झाले आहेत: **{folder_date}**")
            file_map = {f.get("name"): f.get("id") for f in rep_files}
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
                else:
                    xl_rep = pd.ExcelFile(fh)
                    sheet_sel = st.selectbox("शीट (Tab) निवडा", xl_rep.sheet_names)
                    rep_df = pd.read_excel(xl_rep, sheet_name=sheet_sel)

                st.dataframe(
                    rep_df,
                    use_container_width=True,
                    hide_index=True,
                    height=600,
                    column_config=get_smart_col_config(rep_df)
                )

    with tab3:
        st.subheader(f"📊 सर्व Nifty 500 शेअर्सची Live CMP आणि मार्जिन यादी ({len(df_stocks)} शेअर्स)")
        search_q = st.text_input("कंपनी, सिम्बॉल किंवा सेक्टर शोधा (उदा. RELIANCE, BANK, TATA, PHARMA)")
        view_df = df_stocks.copy()
        if search_q:
            view_df = view_df[
                view_df["Symbol"].str.contains(search_q, case=False) |
                view_df["Company"].str.contains(search_q, case=False) |
                view_df["Industry"].str.contains(search_q, case=False)
            ]
        st.dataframe(
            view_df,
            use_container_width=True,
            hide_index=True,
            height=700,
            column_config={
                "Live_CMP": st.column_config.NumberColumn("Live CMP (₹)", format="₹%.2f"),
                "Dhan_MTF": st.column_config.NumberColumn("Dhan MTF %", format="%.2f%%"),
                "Zerodha_MTF": st.column_config.NumberColumn("Zerodha MTF %", format="%.2f%%"),
                "MIS": st.column_config.NumberColumn("MIS Intraday %", format="%.2f%%"),
            }
        )


def add_total_row(df_in, label_col):
    df_clean = df_in[df_in[label_col].astype(str).str.upper() != "TOTAL"].copy()
    tot_c = int(df_clean["Closed Trades"].sum())
    tot_w = int(df_clean["Wins"].sum())
    tot_l = int(df_clean["Losses"].sum())
    tot_wr = round((tot_w / tot_c) * 100, 1) if tot_c > 0 else 0.0
    tot_row = pd.DataFrame([{
        label_col: "TOTAL",
        "Closed Trades": tot_c,
        "Wins": tot_w,
        "Losses": tot_l,
        "Win Rate %": tot_wr,
        "Gross P&L (₹)": round(df_clean["Gross P&L (₹)"].sum(), 2),
        "Charges & Tax (₹)": round(df_clean["Charges & Tax (₹)"].sum(), 2),
        "Net P&L (₹)": round(df_clean["Net P&L (₹)"].sum(), 2),
    }])
    return pd.concat([df_clean, tot_row], ignore_index=True)


def render_tab4_journal(tab4):
    with tab4:
        st.subheader("🏆 एकत्रित ट्रेड जर्नल व इंटरॲक्टिव्ह डॅशबोर्ड")
        if st.button("🔄 ताजा डेटा रिफ्रेश करा"):
            st.cache_data.clear()

        sheets_data, j_err = fetch_master_trade_journal()
        if j_err:
            st.info(j_err)
            return

        if sheets_data and "All_Trades_Journal" in sheets_data:
            all_df = sheets_data["All_Trades_Journal"].copy()

            # खातेधारक आणि ब्रोकर फिल्टर
            f1, f2 = st.columns(2)
            holders = ["सर्व (All)"] + sorted(all_df["ACCOUNT_HOLDER"].dropna().unique().tolist())
            brokers = ["सर्व (All)"] + sorted(all_df["BROKER"].dropna().unique().tolist())
            sel_holder = f1.selectbox("👤 खातेधारक निवडा (Account Holder)", holders)
            sel_broker = f2.selectbox("🏦 ब्रोकर निवडा (Broker Filter)", brokers)

            filtered_df = all_df.copy()
            if sel_holder != "सर्व (All)":
                filtered_df = filtered_df[filtered_df["ACCOUNT_HOLDER"] == sel_holder]
            if sel_broker != "सर्व (All)":
                filtered_df = filtered_df[filtered_df["BROKER"] == sel_broker]

            closed_df = filtered_df[filtered_df["STATUS"].isin(["WIN", "LOSS"])].copy()
            open_cnt = len(filtered_df[filtered_df["STATUS"] == "OPEN"])

            tot_trades = len(closed_df)
            wins = len(closed_df[closed_df["STATUS"] == "WIN"])
            losses = len(closed_df[closed_df["STATUS"] == "LOSS"])
            win_rate = round((wins / tot_trades) * 100, 1) if tot_trades > 0 else 0.0
            net_pnl = round(closed_df["NET_PNL"].sum(), 2) if tot_trades > 0 else 0.0
            tot_charges = round(closed_df["CHARGES_TAX"].sum(), 2) if tot_trades > 0 else 0.0

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("निव्वळ नफा (Net P&L)", f"₹{net_pnl:,.2f}")
            k2.metric("Win Rate %", f"{win_rate}%", f"{wins}W / {losses}L")
            k3.metric("कर व ब्रोकरेज", f"₹{tot_charges:,.2f}")
            k4.metric("चालू ट्रेड्स (OPEN)", f"{open_cnt}")

            if not closed_df.empty:
                closed_df["EXIT_DT"] = pd.to_datetime(closed_df["EXIT_DATE"], errors="coerce")
                closed_df = closed_df.sort_values(by=["EXIT_DT", "TRADE_ID"]).reset_index(drop=True)
                closed_df["Trade_No"] = range(1, len(closed_df) + 1)
                closed_df["Cumulative_PnL"] = closed_df["NET_PNL"].cumsum().round(2)

                # ओळ १: Equity Curve आणि Win/Loss प्रमाण
                g1, g2 = st.columns(2)
                with g1:
                    fig_eq = px.area(
                        closed_df, x="Trade_No", y="Cumulative_PnL",
                        hover_data=["SYMBOL", "ACCOUNT_HOLDER", "BROKER", "NET_PNL"],
                        title="📈 एकूण भांडवल वाढ आलेख (Equity Curve ₹)",
                        color_discrete_sequence=["#00A86B"]
                    )
                    fig_eq.update_layout(height=340, margin=dict(l=10, r=10, t=45, b=10))
                    st.plotly_chart(fig_eq, use_container_width=True)

                with g2:
                    pie_df = pd.DataFrame({"Result": ["Wins", "Losses"], "Count": [wins, losses]})
                    fig_pie = px.pie(
                        pie_df, values="Count", names="Result", hole=0.55,
                        title="🎯 Win vs Loss प्रमाण (Trade Sentiment)",
                        color="Result", color_discrete_map={"Wins": "#00A86B", "Losses": "#E74C3C"}
                    )
                    fig_pie.update_traces(textinfo="value+percent")
                    fig_pie.update_layout(height=340, margin=dict(l=10, r=10, t=45, b=10))
                    st.plotly_chart(fig_pie, use_container_width=True)

                # ओळ २: महिनानिहाय नफा-तोटा आलेख (Monthly P&L Chart)
                valid_m = closed_df.dropna(subset=["EXIT_DT"]).copy()
                if not valid_m.empty:
                    valid_m["YearMonth"] = valid_m["EXIT_DT"].dt.to_period("M")
                    m_grp = valid_m.groupby("YearMonth", as_index=False)["NET_PNL"].sum().sort_values("YearMonth")
                    m_grp["Month"] = m_grp["YearMonth"].dt.strftime("%b-%Y")
                    m_grp["Color"] = m_grp["NET_PNL"].apply(lambda v: "Profit" if v >= 0 else "Loss")
                    m_grp["Label"] = m_grp["NET_PNL"].apply(
                        lambda v: f"₹{v/1000:.1f}K" if abs(v) >= 1000 else f"₹{v:,.0f}"
                    )
                    fig_m = px.bar(
                        m_grp, x="Month", y="NET_PNL", text="Label", color="Color",
                        title="📅 महिनानिहाय नफा-तोटा आलेख (Monthly Net P&L ₹)",
                        color_discrete_map={"Profit": "#00A86B", "Loss": "#E74C3C"}
                    )
                    fig_m.update_traces(textposition="outside")
                    fig_m.update_layout(height=360, showlegend=False, margin=dict(l=10, r=10, t=45, b=10))
                    st.plotly_chart(fig_m, use_container_width=True)
                # ओळ ३: खातेधारक व ब्रोकर तुलना बार चार्ट्स
                g3, g4 = st.columns(2)
                with g3:
                    acc_grp = closed_df.groupby("ACCOUNT_HOLDER", as_index=False)["NET_PNL"].sum()
                    acc_grp["Label"] = acc_grp["NET_PNL"].apply(
                        lambda v: f"₹{v/1000:.1f}K" if abs(v) >= 1000 else f"₹{v:,.0f}"
                    )
                    fig_acc = px.bar(
                        acc_grp, x="ACCOUNT_HOLDER", y="NET_PNL", text="Label",
                        title="👥 खातेधारकानुसार निव्वळ नफा (₹)",
                        color="NET_PNL", color_continuous_scale=["#E74C3C", "#2ECC71"]
                    )
                    fig_acc.update_traces(textposition="outside")
                    fig_acc.update_layout(height=350, showlegend=False, coloraxis_showscale=False, margin=dict(l=10, r=10, t=45, b=10))
                    st.plotly_chart(fig_acc, use_container_width=True)

                with g4:
                    brk_grp = closed_df.groupby("BROKER", as_index=False)["NET_PNL"].sum()
                    brk_grp["Label"] = brk_grp["NET_PNL"].apply(
                        lambda v: f"₹{v/1000:.1f}K" if abs(v) >= 1000 else f"₹{v:,.0f}"
                    )
                    fig_brk = px.bar(
                        brk_grp, x="BROKER", y="NET_PNL", text="Label",
                        title="🏦 ब्रोकरनुसार निव्वळ नफा (₹)",
                        color="BROKER"
                    )
                    fig_brk.update_traces(textposition="outside")
                    fig_brk.update_layout(height=350, showlegend=False, margin=dict(l=10, r=10, t=45, b=10))
                    st.plotly_chart(fig_brk, use_container_width=True)

            # डायनॅमिक सारांश तक्ते तयार करणे
            def build_dynamic_summary(df_sub, group_col, display_label):
                rows = []
                for key, grp in df_sub.groupby(group_col):
                    tot = len(grp)
                    w = len(grp[grp["STATUS"] == "WIN"])
                    l = len(grp[grp["STATUS"] == "LOSS"])
                    wr = round((w / tot) * 100, 1) if tot > 0 else 0.0
                    rows.append({
                        display_label: key,
                        "Closed Trades": tot,
                        "Wins": w,
                        "Losses": l,
                        "Win Rate %": wr,
                        "Gross P&L (₹)": round(grp["GROSS_PNL"].sum(), 2),
                        "Charges & Tax (₹)": round(grp["CHARGES_TAX"].sum(), 2),
                        "Net P&L (₹)": round(grp["NET_PNL"].sum(), 2),
                    })
                return add_total_row(pd.DataFrame(rows), display_label)

            summary_col_cfg = {
                "Win Rate %": st.column_config.NumberColumn(format="%.1f%%"),
                "Gross P&L (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                "Charges & Tax (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                "Net P&L (₹)": st.column_config.NumberColumn(format="₹%.2f"),
            }

            if not closed_df.empty:
                st.subheader("👥 खातेधारकानुसार सारांश (TOTAL सह)")
                st.dataframe(
                    build_dynamic_summary(closed_df, "ACCOUNT_HOLDER", "Account Holder"),
                    use_container_width=True, hide_index=True, column_config=summary_col_cfg
                )

                st.subheader("🏦 ब्रोकरनुसार कामगिरी (TOTAL सह)")
                st.dataframe(
                    build_dynamic_summary(closed_df, "BROKER", "Broker"),
                    use_container_width=True, hide_index=True, column_config=summary_col_cfg
                )

            st.subheader("📋 सर्व ट्रेड्सची यादी (Detailed Trades)")
            st.dataframe(
                filtered_df.drop(columns=["EXIT_DT"], errors="ignore"),
                use_container_width=True,
                hide_index=True,
                height=550,
                column_config={
                    "ENTRY_PRICE": st.column_config.NumberColumn("Entry (₹)", format="₹%.2f"),
                    "EXIT_PRICE": st.column_config.NumberColumn("Exit (₹)", format="₹%.2f"),
                    "GROSS_PNL": st.column_config.NumberColumn("Gross P&L (₹)", format="₹%.2f"),
                    "CHARGES_TAX": st.column_config.NumberColumn("Charges (₹)", format="₹%.2f"),
                    "NET_PNL": st.column_config.NumberColumn("Net P&L (₹)", format="₹%.2f"),
                    "CUMULATIVE_PNL": st.column_config.NumberColumn("Cumulative P&L (₹)", format="₹%.2f"),
                    "RETURN_%": st.column_config.NumberColumn("Return %", format="%.2f%%"),
                }
            )


def render_app():
    st.title("📈 Nifty 500 MTF & MIS Pro")
    st.caption("Dhan, Zerodha, FYERS, Groww — पोझिशन सायझिंग, ब्रेकआऊट रिपोर्ट्स आणि ट्रेड जर्नल")

    df_stocks = load_stocks_and_margins()
    if df_stocks.empty:
        st.error("Nifty_500_Stocks.xlsx सापडली नाही! कृपया GitHub रेपॉझिटरी तपासा.")
        return

    tab1, tab2, tab3, tab4 = st.tabs([
        "🧮 कॅल्क्युलेटर",
        "🚀 ब्रेकआऊट रिपोर्ट्स",
        "📊 ५०० शेअर्स मार्जिन",
        "🏆 ट्रेड जर्नल (P&L)"
    ])

    with tab1:
        render_tab1_calculator(df_stocks)
    render_tab2_and_tab3(tab2, tab3, df_stocks)
    render_tab4_journal(tab4)


if __name__ == "__main__":
    render_app()
