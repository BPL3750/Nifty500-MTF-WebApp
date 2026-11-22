
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

import re
from datetime import datetime


@st.cache_data(ttl=900)
def get_all_report_date_folders():
    service, root_id, err = get_gdrive_service()
    if err:
        return [], err
    try:
        q_daily = f"name='01_Daily_Reports' and '{root_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        daily_res = service.files().list(q=q_daily, fields="files(id)").execute().get("files", [])
        if not daily_res:
            return [], "Drive वर '01_Daily_Reports' फोल्डर सापडले नाही."

        daily_id = next(iter(daily_res)).get("id")
        q_dates = f"'{daily_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        date_folders = service.files().list(
            q=q_dates, orderBy="createdTime desc", pageSize=60, fields="files(id, name)"
        ).execute().get("files", [])

        parsed_list = []
        for f in date_folders:
            fname = f.get("name", "").strip()
            try:
                dt = datetime.strptime(fname, "%d-%b-%Y")
                parsed_list.append({
                    "id": f.get("id"),
                    "name": fname,
                    "dt": dt,
                    "month": dt.strftime("%b-%Y"),
                    "day": dt.strftime("%A"),
                })
            except ValueError:
                continue

        parsed_list.sort(key=lambda x: x["dt"], reverse=True)
        return parsed_list, None
    except Exception as e:
        return [], f"Drive फोल्डर वाचताना त्रुटी: {e}"


@st.cache_data(ttl=900)
def fetch_clean_reports_for_date(date_folder_id):
    """एकाच दिवसातील डुप्लिकेट रिपोर्ट्स वगळून प्रत्येक रिपोर्टचा फक्त ताजा (Latest) व्हर्जन आणणे"""
    service, _, err = get_gdrive_service()
    if err:
        return [], None, err

    try:
        folder_ids = [date_folder_id]
        q_subs = f"'{date_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        sub_folders = service.files().list(q=q_subs, fields="files(id, name)").execute().get("files", [])
        for sf in sub_folders:
            folder_ids.append(sf.get("id"))

        raw_files = []
        for fid in folder_ids:
            q_files = f"'{fid}' in parents and trashed=false and (name contains '.xlsx' or name contains '.csv')"
            items = service.files().list(
                q=q_files, orderBy="createdTime desc", pageSize=50, fields="files(id, name, createdTime)"
            ).execute().get("files", [])
            for it in items:
                if "Calculator" not in it.get("name", ""):
                    raw_files.append(it)

        raw_files.sort(key=lambda x: x.get("createdTime", ""), reverse=True)

        unique_reports = {}
        for it in raw_files:
            fname = it.get("name", "")
            # फाईलच्या नावातील तारीख व वेळ काढून मूळ रिपोर्टचे नाव मिळवणे (उदा. Breakout_Scanner.xlsx)
            clean_base = re.sub(r"_\d{1,2}-[A-Za-z]{3}-\d{2,4}(_\d{4})?", "", fname)
            time_match = re.search(r"_(\d{2})(\d{2})\.(xlsx|csv)$", fname, re.IGNORECASE)
            time_tag = f" (Latest: {time_match.group(1)}:{time_match.group(2)})" if time_match else ""
            display_label = f"{clean_base}{time_tag}"

            if clean_base not in unique_reports:
                unique_reports[clean_base] = {
                    "id": it.get("id"),
                    "name": fname,
                    "label": display_label
                }

        return list(unique_reports.values()), service, None
    except Exception as e:
        return [], None, f"रिपोर्ट्स वाचताना त्रुटी: {e}"


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
        st.subheader("🚀 Google Drive वरील ब्रेकआऊट आणि सेक्टर रिपोर्ट्स")
        date_items, f_err = get_all_report_date_folders()

        if f_err:
            st.info(f_err)
        elif not date_items:
            st.warning("अजून कोणतेही तारीख फोल्डर तयार झालेले नाही.")
        else:
            # १. महिना (Month), वार (Day) आणि तारीख (Date) असे ३ स्वतंत्र ड्रॉपडाउन
            months = list(dict.fromkeys([d["month"] for d in date_items]))
            c_m, c_day, c_dt = st.columns(3)

            sel_month = c_m.selectbox("📅 महिना निवडा (Month)", months, key="rep_month")
            month_filtered = [d for d in date_items if d["month"] == sel_month]

            days_in_month = ["सर्व वार (All Days)"] + list(dict.fromkeys([d["day"] for d in month_filtered]))
            sel_day = c_day.selectbox("🗓️ वार निवडा (Day)", days_in_month, key="rep_day")

            final_dates = month_filtered if sel_day == "सर्व वार (All Days)" else [d for d in month_filtered if d["day"] == sel_day]
            date_map = {f"{d['name']} ({d['day']})": d["id"] for d in final_dates}

            sel_date_label = c_dt.selectbox("📆 तारीख निवडा (Date)", list(date_map.keys()), key="rep_date")
            selected_folder_id = date_map[sel_date_label]

            # २. निवडलेल्या तारखेचे स्वच्छ (Non-duplicate) रिपोर्ट्स आणणे
            clean_reps, service_obj, r_err = fetch_clean_reports_for_date(selected_folder_id)
            if r_err:
                st.error(r_err)
            elif not clean_reps:
                st.warning(f"📅 **{sel_date_label}** या दिवशी कोणताही रिपोर्ट सापडला नाही.")
            else:
                st.success(f"📅 लोड झालेली तारीख: **{sel_date_label}** ({len(clean_reps)} युनिक रिपोर्ट्स)")
                rep_map = {r["label"]: r for r in clean_reps}
                chosen_label = st.selectbox("📑 रिपोर्ट निवडा (फक्त ताजा व्हर्जन)", list(rep_map.keys()), key="rep_file")

                if chosen_label:
                    chosen_rep = rep_map[chosen_label]
                    request = service_obj.files().get_media(fileId=chosen_rep["id"])
                    fh = io.BytesIO()
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while not done:
                        _, done = downloader.next_chunk()
                    fh.seek(0)

                    if chosen_rep["name"].lower().endswith(".csv"):
                        rep_df = pd.read_csv(fh)
                    else:
                        xl_rep = pd.ExcelFile(fh)
                        sheet_sel = st.selectbox("शीट (Tab) निवडा", xl_rep.sheet_names, key="rep_sheet")
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


@st.cache_data(ttl=900)
def get_daily_date_folders():
    service, root_id, err = get_gdrive_service()
    if err:
        return [], err
    try:
        q_daily = f"name='01_Daily_Reports' and '{root_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        daily_res = service.files().list(q=q_daily, fields="files(id)").execute().get("files", [])
        if not daily_res:
            return [], "Drive वर '01_Daily_Reports' फोल्डर सापडले नाही."
        daily_id = next(iter(daily_res)).get("id")
        q_dates = f"'{daily_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        folders = service.files().list(
            q=q_dates, orderBy="createdTime desc", pageSize=15, fields="files(id, name)"
        ).execute().get("files", [])
        return folders, None
    except Exception as e:
        return [], f"तारीख फोल्डर वाचताना त्रुटी: {e}"


def find_subfolder_id(service, parent_id, folder_name):
    q = f"name='{folder_name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    res = service.files().list(q=q, fields="files(id)").execute().get("files", [])
    return next(iter(res)).get("id") if res else None


@st.cache_data(ttl=900)
def fetch_charts_and_excel(date_folder_id, mode):
    service, _, err = get_gdrive_service()
    if err:
        return pd.DataFrame(), {}, err

    try:
        if mode == "breakout":
            scan_dir_id = find_subfolder_id(service, date_folder_id, "Full_Market_Scan")
            excel_parent = scan_dir_id or date_folder_id
            excel_prefix = "Breakout_Scanner"
            chart_dir_id = find_subfolder_id(service, scan_dir_id, "Charts") if scan_dir_id else None
        else:
            excel_parent = date_folder_id
            excel_prefix = "Analyze_Trending_charts"
            chart_dir_id = find_subfolder_id(service, date_folder_id, "Trendline_Charts")

        q_xl = f"'{excel_parent}' in parents and name contains '{excel_prefix}' and name contains '.xlsx' and trashed=false"
        xl_files = service.files().list(
            q=q_xl, orderBy="createdTime desc", pageSize=5, fields="files(id, name)"
        ).execute().get("files", [])

        df_excel = pd.DataFrame()
        if xl_files:
            latest_xl = next(iter(xl_files))
            req = service.files().get_media(fileId=latest_xl.get("id"))
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.seek(0)
            xl_obj = pd.ExcelFile(fh)
            df_excel = pd.read_excel(xl_obj, sheet_name=next(iter(xl_obj.sheet_names)))

        chart_map = {}
        if chart_dir_id:
            q_png = f"'{chart_dir_id}' in parents and name contains '.png' and trashed=false"
            png_files = service.files().list(
                q=q_png, orderBy="createdTime desc", pageSize=250, fields="files(id, name)"
            ).execute().get("files", [])

            for pf in png_files:
                fname = pf.get("name", "")
                sym = fname.split("_").pop(0).strip().upper()
                if sym and sym not in chart_map:
                    chart_map[sym] = {"id": pf.get("id"), "name": fname}

        return df_excel, chart_map, None
    except Exception as e:
        return pd.DataFrame(), {}, f"चार्ट डेटा वाचताना त्रुटी: {e}"


@st.cache_data(ttl=1800)
def download_chart_bytes(file_id):
    service, _, err = get_gdrive_service()
    if err:
        return None
    req = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return fh.getvalue()


def render_chart_scanner_tab(tab_obj, mode, tab_heading):
    with tab_obj:
        st.subheader(tab_heading)
        folders, f_err = get_daily_date_folders()
        if f_err:
            st.info(f_err)
            return
        if not folders:
            st.warning("अजून कोणतेही तारीख फोल्डर सापडले नाही.")
            return

        folder_map = {f.get("name"): f.get("id") for f in folders}
        c_date, c_view = st.columns((1, 2))
        sel_date = c_date.selectbox("📅 तारीख निवडा (Date)", list(folder_map.keys()), key=f"dt_{mode}")
        view_mode = c_view.radio(
            "दृश्य निवडा (View Mode)",
            ["🔍 एक शेअर निवडा (Single Chart + Data)", "🖼️ सर्व चार्ट्स गॅलरी (All Charts Grid)"],
            horizontal=True,
            key=f"vm_{mode}"
        )

        df_excel, chart_map, c_err = fetch_charts_and_excel(folder_map[sel_date], mode)
        if c_err:
            st.error(c_err)
            return

        excel_syms = []
        sym_col = None
        for col in df_excel.columns:
            if str(col).strip().upper() in ("SYMBOL", "STOCK", "TICKER"):
                sym_col = col
                for s in df_excel[col].dropna().unique():
                    s_clean = str(s).strip().upper()
                    if s_clean and "STOCK" not in s_clean and "TOTAL" not in s_clean and "PORTFOLIO" not in s_clean:
                        excel_syms.append(s_clean)
                break

        chart_syms = sorted(list(chart_map.keys()))
        only_data_syms = sorted([s for s in set(excel_syms) if s not in chart_map])
        all_syms = chart_syms + only_data_syms

        if not all_syms and df_excel.empty:
            st.warning(f"📅 **{sel_date}** या तारखेसाठी चार्ट्स किंवा रिपोर्ट उपलब्ध नाहीत.")
            return

        if view_mode.startswith("🔍"):
            if all_syms:
                sel_sym = st.selectbox(
                    f"📈 शेअर निवडा ({len(chart_syms)} चार्ट्स उपलब्ध | एकूण {len(all_syms)} शेअर्स)",
                    all_syms,
                    key=f"sym_{mode}"
                )

                if not df_excel.empty and sym_col:
                    match_rows = df_excel[df_excel[sym_col].astype(str).str.strip().str.upper() == sel_sym]
                    if not match_rows.empty:
                        st.markdown(f"**📋 {sel_sym} — स्कॅनर विश्लेषण तपशील:**")
                        st.dataframe(
                            match_rows,
                            use_container_width=True,
                            hide_index=True,
                            column_config=get_smart_col_config(match_rows)
                        )

                if sel_sym in chart_map:
                    img_bytes = download_chart_bytes(chart_map[sel_sym]["id"])
                    if img_bytes:
                        st.image(img_bytes, caption=f"{sel_sym} — {chart_map[sel_sym]['name']}", use_container_width=True)
                else:
                    st.info(f"ℹ️ **{sel_sym}** साठी PNG चार्ट इमेज उपलब्ध नाही.")
        else:
            if not chart_map:
                st.info("या तारखेसाठी PNG चार्ट्स सापडले नाहीत.")
            else:
                st.success(f"एकूण **{len(chart_map)}** ताजे चार्ट्स लोड झाले आहेत:")
                grid_cols = st.columns(2)
                for idx, (sym, info) in enumerate(sorted(chart_map.items())):
                    img_bytes = download_chart_bytes(info["id"])
                    if img_bytes:
                        with grid_cols[idx % 2]:
                            with st.container(border=True):
                                st.image(img_bytes, caption=f"{sym} ({info['name']})", use_container_width=True)

        if not df_excel.empty:
            with st.expander(f"📊 संपूर्ण Excel डेटा तक्ता पहा ({len(df_excel)} ओळी)", expanded=False):
                st.dataframe(
                    df_excel,
                    use_container_width=True,
                    hide_index=True,
                    height=500,
                    column_config=get_smart_col_config(df_excel)
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

def render_journal_charts(closed_df, wins, losses):
    closed_df = closed_df.copy()
    closed_df["EXIT_DT"] = pd.to_datetime(closed_df["EXIT_DATE"], errors="coerce")
    closed_df = closed_df.sort_values(by=["EXIT_DT", "TRADE_ID"]).reset_index(drop=True)
    closed_df["Trade_No"] = range(1, len(closed_df) + 1)
    closed_df["Cumulative_PnL"] = closed_df["NET_PNL"].cumsum().round(2)

    g1, g2 = st.columns(2)
    with g1:
        with st.container(border=True):
            fig_eq = px.area(
                closed_df, x="Trade_No", y="Cumulative_PnL",
                hover_data=["SYMBOL", "ACCOUNT_HOLDER", "BROKER", "SEGMENT", "NET_PNL"],
                title="📈 एकूण भांडवल वाढ आलेख (Equity Curve ₹)",
                color_discrete_sequence=["#00A86B"]
            )
            fig_eq.update_layout(height=340, margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig_eq, use_container_width=True)

    with g2:
        with st.container(border=True):
            pie_df = pd.DataFrame({"Result": ["Wins", "Losses"], "Count": [wins, losses]})
            fig_pie = px.pie(
                pie_df, values="Count", names="Result", hole=0.55,
                title="🎯 Win vs Loss प्रमाण (Trade Sentiment)",
                color="Result", color_discrete_map={"Wins": "#00A86B", "Losses": "#E74C3C"}
            )
            fig_pie.update_traces(textinfo="value+percent")
            fig_pie.update_layout(height=340, margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig_pie, use_container_width=True)

    valid_m = closed_df.dropna(subset=["EXIT_DT"]).copy()
    if not valid_m.empty:
        with st.container(border=True):
            valid_m["YearMonth"] = valid_m["EXIT_DT"].dt.to_period("M")
            m_grp = valid_m.groupby("YearMonth", as_index=False)["NET_PNL"].sum().sort_values("YearMonth")
            m_grp["Month"] = m_grp["YearMonth"].dt.strftime("%b-%Y")
            m_grp["Color"] = m_grp["NET_PNL"].apply(lambda v: "Profit" if v >= 0 else "Loss")
            m_grp["Label"] = m_grp["NET_PNL"].apply(lambda v: f"₹{v/1000:.1f}K" if abs(v) >= 1000 else f"₹{v:,.0f}")
            fig_m = px.bar(
                m_grp, x="Month", y="NET_PNL", text="Label", color="Color",
                title="📅 महिनानिहाय नफा-तोटा आलेख (Monthly Net P&L ₹)",
                color_discrete_map={"Profit": "#00A86B", "Loss": "#E74C3C"}
            )
            fig_m.update_traces(textposition="outside")
            fig_m.update_layout(height=360, showlegend=False, margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig_m, use_container_width=True)

    g3, g4 = st.columns(2)
    with g3:
        with st.container(border=True):
            acc_grp = closed_df.groupby("ACCOUNT_HOLDER", as_index=False)["NET_PNL"].sum()
            acc_grp["Label"] = acc_grp["NET_PNL"].apply(lambda v: f"₹{v/1000:.1f}K" if abs(v) >= 1000 else f"₹{v:,.0f}")
            fig_acc = px.bar(
                acc_grp, x="ACCOUNT_HOLDER", y="NET_PNL", text="Label",
                title="👥 खातेधारकानुसार निव्वळ नफा (₹)",
                color="NET_PNL", color_continuous_scale=["#E74C3C", "#2ECC71"]
            )
            fig_acc.update_traces(textposition="outside")
            fig_acc.update_layout(height=350, showlegend=False, coloraxis_showscale=False, margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig_acc, use_container_width=True)

    with g4:
        with st.container(border=True):
            brk_grp = closed_df.groupby("BROKER", as_index=False)["NET_PNL"].sum()
            brk_grp["Label"] = brk_grp["NET_PNL"].apply(lambda v: f"₹{v/1000:.1f}K" if abs(v) >= 1000 else f"₹{v:,.0f}")
            fig_brk = px.bar(
                brk_grp, x="BROKER", y="NET_PNL", text="Label",
                title="🏦 ब्रोकरनुसार निव्वळ नफा (₹)",
                color="BROKER"
            )
            fig_brk.update_traces(textposition="outside")
            fig_brk.update_layout(height=350, showlegend=False, margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig_brk, use_container_width=True)


def render_tab4_journal(tab4):
    with tab4:
        j_col1, j_col2 = st.columns((5, 1))
        j_col1.subheader("🏆 एकत्रित ट्रेड जर्नल व इंटरॲक्टिव्ह डॅशबोर्ड")

        app_pin = str(st.secrets.get("JOURNAL_PASSWORD", "1234")).strip()
        if "journal_unlocked" not in st.session_state:
            st.session_state["journal_unlocked"] = False

        if not st.session_state["journal_unlocked"]:
            st.info("🔒 हा विभाग खाजगी आहे. सर्व खातेधारकांचा नफा-तोटा पाहण्यासाठी पासवर्ड/PIN टाका.")
            with st.form("journal_login_form"):
                entered_pin = st.text_input("पासवर्ड / PIN टाका", type="password")
                submitted = st.form_submit_button("🔓 अनलॉक करा (Unlock)")
                if submitted:
                    if entered_pin.strip() == app_pin:
                        st.session_state["journal_unlocked"] = True
                        st.rerun()
                    else:
                        st.error("❌ चुकीचा पासवर्ड! Streamlit Secrets मधील JOURNAL_PASSWORD तपासा.")
            return


        if j_col2.button("🔒 लॉक करा", key="lock_journal_btn", use_container_width=True):
            st.session_state["journal_unlocked"] = False
            st.rerun()

        sheets_data, j_err = fetch_master_trade_journal()
        if j_err:
            st.info(j_err)
            return

        if sheets_data and "All_Trades_Journal" in sheets_data:
            all_df = sheets_data["All_Trades_Journal"].copy()

            f1, f2, f3 = st.columns(3)
            holders = ["सर्व (All)"] + sorted(all_df["ACCOUNT_HOLDER"].dropna().unique().tolist())
            brokers = ["सर्व (All)"] + sorted(all_df["BROKER"].dropna().unique().tolist())
            segments = ["सर्व (All)"] + sorted(all_df["SEGMENT"].dropna().unique().tolist())

            sel_holder = f1.selectbox("👤 खातेधारक (Account Holder)", holders)
            sel_broker = f2.selectbox("🏦 ब्रोकर (Broker)", brokers)
            sel_seg = f3.selectbox("📊 सेगमेंट (Intraday / Swing)", segments)

            filtered_df = all_df.copy()
            if sel_holder != "सर्व (All)":
                filtered_df = filtered_df[filtered_df["ACCOUNT_HOLDER"] == sel_holder]
            if sel_broker != "सर्व (All)":
                filtered_df = filtered_df[filtered_df["BROKER"] == sel_broker]
            if sel_seg != "सर्व (All)":
                filtered_df = filtered_df[filtered_df["SEGMENT"] == sel_seg]

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
                render_journal_charts(closed_df, wins, losses)

                summary_col_cfg = {
                    "Win Rate %": st.column_config.NumberColumn(format="%.1f%%"),
                    "Gross P&L (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                    "Charges & Tax (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                    "Net P&L (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                }

                st.subheader("⚡ Intraday विरुद्ध Swing कामगिरी तुलना (TOTAL सह)")
                st.dataframe(build_dynamic_summary(closed_df, "SEGMENT", "Segment"), use_container_width=True, hide_index=True, column_config=summary_col_cfg)

                sym_rows = []
                for sym, grp in closed_df.groupby("SYMBOL"):
                    tot = len(grp)
                    w = len(grp[grp["STATUS"] == "WIN"])
                    wr = round((w / tot) * 100, 1) if tot > 0 else 0.0
                    comp_name = grp["COMPANY_NAME"].dropna().iloc.__getitem__(0) if "COMPANY_NAME" in grp.columns and not grp["COMPANY_NAME"].dropna().empty else sym
                    sym_rows.append({"Symbol": sym, "Company": comp_name, "Trades": tot, "Win Rate %": wr, "Net P&L (₹)": round(grp["NET_PNL"].sum(), 2)})
                sym_df = pd.DataFrame(sym_rows)
                top_winners = sym_df[sym_df["Net P&L (₹)"] > 0].sort_values(by="Net P&L (₹)", ascending=False).head(5)
                top_losers = sym_df[sym_df["Net P&L (₹)"] < 0].sort_values(by="Net P&L (₹)", ascending=True).head(5)

                stock_col_cfg = {
                    "Symbol": st.column_config.TextColumn("Symbol", width="small"),
                    "Company": st.column_config.TextColumn("Company Name", width="large"),
                    "Trades": st.column_config.NumberColumn("Trades", width="small"),
                    "Win Rate %": st.column_config.NumberColumn("Win %", format="%.1f%%", width="small"),
                    "Net P&L (₹)": st.column_config.NumberColumn("Net P&L (₹)", format="₹%.2f", width="medium"),
                }

                st.subheader("🏅 सर्वाधिक नफा व तोटा देणारे टॉप ५ शेअर्स")
                t_win_col, t_loss_col = st.columns(2)
                with t_win_col:
                    st.markdown("**🟢 टॉप ५ नफा देणारे शेअर्स (Top 5 Winners)**")
                    if not top_winners.empty:
                        st.dataframe(top_winners, use_container_width=True, hide_index=True, height=(len(top_winners) + 1) * 35 + 5, column_config=stock_col_cfg)
                    else:
                        st.info("नफा झालेले शेअर्स नाहीत.")

                with t_loss_col:
                    st.markdown("**🔴 टॉप ५ तोटा देणारे शेअर्स (Top 5 Losers)**")
                    if not top_losers.empty:
                        st.dataframe(top_losers, use_container_width=True, hide_index=True, height=(len(top_losers) + 1) * 35 + 5, column_config=stock_col_cfg)
                    else:
                        st.info("तोटा झालेले शेअर्स नाहीत.")

                st.subheader("👥 खातेधारकानुसार सारांश (TOTAL सह)")
                st.dataframe(build_dynamic_summary(closed_df, "ACCOUNT_HOLDER", "Account Holder"), use_container_width=True, hide_index=True, column_config=summary_col_cfg)

                st.subheader("🏦 ब्रोकरनुसार कामगिरी (TOTAL सह)")
                st.dataframe(build_dynamic_summary(closed_df, "BROKER", "Broker"), use_container_width=True, hide_index=True, column_config=summary_col_cfg)

            st.subheader("📋 सर्व ट्रेड्सची यादी (Detailed Trades)")
            st.dataframe(
                filtered_df,
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

@st.cache_data(ttl=1800)
def build_price_volume_matrix(symbols_tuple, company_map, industry_map):
    """२०२५ पासूनचा सरासरी व्हॉल्यूम, २०-दिवस व्हॉल्यूम आणि किमतीतील बदलाचा डेटा तयार करणे"""
    rows = []
    try:
        tickers = [f"{s}.NS" for s in symbols_tuple if s and not s.startswith("DUMMY")]
        data = yf.download(tickers, start="2025-01-01", progress=False)
        if data.empty or "Close" not in data or "Volume" not in data:
            return pd.DataFrame()

        close_df = data["Close"]
        vol_df = data["Volume"]

        for s in symbols_tuple:
            col = f"{s}.NS"
            if col not in close_df.columns or col not in vol_df.columns:
                continue

            c_ser = close_df[col].dropna()
            v_ser = vol_df[col].dropna()
            if len(c_ser) < 22 or len(v_ser) < 22:
                continue

            cmp = round(float(c_ser.iloc[-1]), 2)
            prev_1d = float(c_ser.iloc[-2])
            prev_5d = float(c_ser.iloc[-6]) if len(c_ser) >= 6 else prev_1d

            chg_1d = round(((cmp - prev_1d) / prev_1d) * 100, 2) if prev_1d > 0 else 0.0
            chg_5d = round(((cmp - prev_5d) / prev_5d) * 100, 2) if prev_5d > 0 else 0.0

            curr_vol = int(v_ser.iloc[-1])
            avg_vol_20d = int(v_ser.iloc[-21:-1].mean())
            avg_vol_2025 = int(v_ser.iloc[:-1].mean())

            vol_ratio_20d = round(curr_vol / avg_vol_20d, 2) if avg_vol_20d > 0 else 0.0
            vol_ratio_2025 = round(curr_vol / avg_vol_2025, 2) if avg_vol_2025 > 0 else 0.0

            rows.append({
                "Symbol": s,
                "Company": company_map.get(s, s),
                "Sector": industry_map.get(s, "-"),
                "CMP (₹)": cmp,
                "1D Chg %": chg_1d,
                "5D Chg %": chg_5d,
                "Current Vol": curr_vol,
                "20D Avg Vol": avg_vol_20d,
                "2025+ Avg Vol": avg_vol_2025,
                "Vol vs 20D (x)": vol_ratio_20d,
                "Vol vs 2025+ (x)": vol_ratio_2025,
            })
    except Exception:
        pass

    return pd.DataFrame(rows)

def render_price_volume_tab(tab_obj, df_stocks):
    with tab_obj:
        st.subheader("📊 Nifty 500 प्राईस-व्हॉल्यूम मॅट्रिक्स (2025+ Baseline)")
        st.caption("२०२५ पासूनच्या सरासरी व्हॉल्यूम आणि किमतीतील बदलावर आधारित ४ प्रमुख मार्केट संकेत")

        c1, c2, c3 = st.columns((1.2, 1.2, 1.6))
        tf_mode = c1.radio("कालावधी (Timeframe)", ["१ दिवस (1D Change)", "५ दिवस (5D Swing)"], horizontal=True)
        base_mode = c2.radio("व्हॉल्यूम तुलना (Baseline)", ["2025+ Avg Vol", "20D Avg Vol"], horizontal=True)
        vol_thresh = c3.slider("High Volume मर्यादा (किमान पट / x)", min_value=1.2, max_value=3.0, value=1.5, step=0.1)

        comp_map = dict(zip(df_stocks["Symbol"], df_stocks["Company"]))
        ind_map = dict(zip(df_stocks["Symbol"], df_stocks["Industry"]))

        with st.spinner("Nifty 500 शेअर्सचा २०२५ पासूनचा व्हॉल्यूम व प्राईस डेटा तपासत आहे..."):
            pv_df = build_price_volume_matrix(tuple(df_stocks["Symbol"].tolist()), comp_map, ind_map)

        if pv_df.empty:
            st.warning("सध्या प्राईस-व्हॉल्यूम डेटा लोड होऊ शकला नाही. कृपया थोड्या वेळाने रिफ्रेश करा.")
            return

        chg_col = "1D Chg %" if tf_mode.startswith("१") else "5D Chg %"
        ratio_col = "Vol vs 2025+ (x)" if base_mode.startswith("2025") else "Vol vs 20D (x)"

        def classify_quadrant(r):
            v_rat = r[ratio_col]
            p_chg = r[chg_col]
            if v_rat >= vol_thresh and p_chg > 0:
                return "1. Strong Uptrend / Breakout"
            elif v_rat >= vol_thresh and p_chg < 0:
                return "2. Strong Downtrend / Panic Selling"
            elif v_rat < 1.0 and p_chg >= 1.0:
                return "3. Weak Rally / Trap"
            elif v_rat < 1.0 and p_chg <= -1.0:
                return "4. Lack of Interest / Pullback"
            return "Neutral"

        pv_df["Signal_Category"] = pv_df.apply(classify_quadrant, axis=1)

        q1_df = pv_df[pv_df["Signal_Category"].str.startswith("1.")].sort_values(by=ratio_col, ascending=False)
        q2_df = pv_df[pv_df["Signal_Category"].str.startswith("2.")].sort_values(by=ratio_col, ascending=False)
        q3_df = pv_df[pv_df["Signal_Category"].str.startswith("3.")].sort_values(by=chg_col, ascending=False)
        q4_df = pv_df[pv_df["Signal_Category"].str.startswith("4.")].sort_values(by=chg_col, ascending=True)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("🟢 1. Strong Breakout", f"{len(q1_df)} शेअर्स", f"Vol ≥ {vol_thresh}x & Price +")
        m2.metric("🔴 2. Panic Selling", f"{len(q2_df)} शेअर्स", f"Vol ≥ {vol_thresh}x & Price -")
        m3.metric("🟡 3. Weak Rally / Trap", f"{len(q3_df)} शेअर्स", "Low Vol & Price +")
        m4.metric("🔵 4. Pullback / Low Interest", f"{len(q4_df)} शेअर्स", "Low Vol & Price -")

        # यादी निवडण्याचा ड्रॉपडाउन
        sel_cat = st.selectbox(
            "🔍 कोणती यादी पाहायची आहे ते निवडा:",
            [
                f"🟢 1. Strong Uptrend / Breakout ({len(q1_df)} शेअर्स — खरेदीसाठी मजबूत संकेत)",
                f"🔴 2. Strong Downtrend / Panic Selling ({len(q2_df)} शेअर्स — विक्रीचा मोठा दबाव)",
                f"🟡 3. Weak Rally / Trap ({len(q3_df)} शेअर्स — फसवी तेजी, सावध राहावे)",
                f"🔵 4. Lack of Interest / Pullback ({len(q4_df)} शेअर्स — तात्पुरती घसरण)",
                f"📋 सर्व Nifty 500 शेअर्स ({len(pv_df)} शेअर्स)"
            ]
        )

        if sel_cat.startswith("🟢"):
            show_df = q1_df
            bar_color = "#00A86B"
            clean_file_prefix = "1_Strong_Uptrend_Breakout"
            chart_title = f"🟢 Top 15 Breakout शेअर्स ({base_mode})"
        elif sel_cat.startswith("🔴"):
            show_df = q2_df
            bar_color = "#E74C3C"
            clean_file_prefix = "2_Strong_Downtrend_Panic_Selling"
            chart_title = f"🔴 Top 15 Panic Selling शेअर्स ({base_mode})"
        elif sel_cat.startswith("🟡"):
            show_df = q3_df
            bar_color = "#F39C12"
            clean_file_prefix = "3_Weak_Rally_Trap"
            chart_title = "🟡 Top 15 Weak Rally / Trap शेअर्स (कमी व्हॉल्यूम तेजी)"
        elif sel_cat.startswith("🔵"):
            show_df = q4_df
            bar_color = "#3498DB"
            clean_file_prefix = "4_Lack_of_Interest_Pullback"
            chart_title = "🔵 Top 15 Pullback शेअर्स (कमी व्हॉल्यूम घसरण)"
        else:
            show_df = pv_df.sort_values(by=ratio_col, ascending=False)
            bar_color = "#1F4E78"
            clean_file_prefix = "All_Nifty500_Price_Volume"
            chart_title = f"📋 सर्वाधिक व्हॉल्यूम असलेले Top 15 शेअर्स ({base_mode})"

        # २ स्पष्ट आणि अर्थपूर्ण बार चार्ट्स
        g_left, g_right = st.columns((1.2, 1))
        with g_left:
            with st.container(border=True):
                if not show_df.empty:
                    top_bar = show_df.head(15).copy()
                    x_metric = chg_col if sel_cat.startswith(("🟡", "🔵")) else ratio_col
                    top_bar["Bar_Label"] = top_bar.apply(
                        lambda r: f" {r[ratio_col]:.2f}x | {r[chg_col]:+.1f}%", axis=1
                    )
                    top_bar = top_bar.iloc[::-1]
                    fig_bar = px.bar(
                        top_bar,
                        x=x_metric,
                        y="Symbol",
                        orientation="h",
                        text="Bar_Label",
                        hover_data=["Company", "Sector", "CMP (₹)", "Current Vol", "2025+ Avg Vol"],
                        title=chart_title,
                        color_discrete_sequence=[bar_color]
                    )
                    fig_bar.update_traces(textposition="outside", cliponaxis=False)
                    fig_bar.update_layout(
                        height=440,
                        yaxis_title="",
                        xaxis_title=x_metric,
                        margin=dict(l=10, r=70, t=45, b=10)
                    )
                    st.plotly_chart(fig_bar, use_container_width=True)
                else:
                    st.info("या श्रेणीत सध्या कोणताही शेअर नाही.")

        with g_right:
            with st.container(border=True):
                sec_sub = pv_df[pv_df["Signal_Category"].str.startswith(("1.", "2."))].copy()
                if not sec_sub.empty:
                    sec_grp = sec_sub.groupby(["Sector", "Signal_Category"], as_index=False).size()
                    top_secs = sec_sub["Sector"].value_counts().head(8).index.tolist()
                    sec_grp = sec_grp[sec_grp["Sector"].isin(top_secs)]
                    fig_sec = px.bar(
                        sec_grp,
                        x="size",
                        y="Sector",
                        color="Signal_Category",
                        orientation="h",
                        barmode="group",
                        text="size",
                        title="🏭 सेक्टरनुसार Breakout विरुद्ध Panic Selling",
                        color_discrete_map={
                            "1. Strong Uptrend / Breakout": "#00A86B",
                            "2. Strong Downtrend / Panic Selling": "#E74C3C"
                        }
                    )
                    fig_sec.update_traces(textposition="outside")
                    fig_sec.update_layout(
                        height=440,
                        yaxis_title="",
                        xaxis_title="शेअर्सची संख्या",
                        legend_title="",
                        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
                        margin=dict(l=10, r=30, t=65, b=10)
                    )
                    st.plotly_chart(fig_sec, use_container_width=True)
                else:
                    st.info("सेक्टर तुलना डेटा उपलब्ध नाही.")

        # अचूक नावासह थेट CSV डाउनलोड बटन
        export_df = show_df.drop(columns=["Signal_Category"], errors="ignore")
        today_str = datetime.now().strftime("%d-%b-%Y")
        dl_filename = f"{clean_file_prefix}_{len(export_df)}_Stocks_{today_str}.csv"

        dl_c1, dl_c2 = st.columns((3, 1.2))
        dl_c1.markdown(f"**📋 निवडलेली यादी ({len(export_df)} शेअर्स):**")
        dl_c2.download_button(
            label="📥 ही यादी CSV डाउनलोड करा",
            data=export_df.to_csv(index=False).encode("utf-8-sig"),
            file_name=dl_filename,
            mime="text/csv",
            use_container_width=True
        )

        st.dataframe(
            export_df,
            use_container_width=True,
            hide_index=True,
            height=550,
            column_config={
                "CMP (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                "1D Chg %": st.column_config.NumberColumn(format="%.2f%%"),
                "5D Chg %": st.column_config.NumberColumn(format="%.2f%%"),
                "Current Vol": st.column_config.NumberColumn(format="%d"),
                "20D Avg Vol": st.column_config.NumberColumn(format="%d"),
                "2025+ Avg Vol": st.column_config.NumberColumn(format="%d"),
                "Vol vs 20D (x)": st.column_config.NumberColumn(format="%.2fx"),
                "Vol vs 2025+ (x)": st.column_config.NumberColumn(format="%.2fx"),
            }
        )



def render_app():
    top_c1, top_c2 = st.columns((5, 1))
    with top_c1:
        st.title("📈 Nifty 500 MTF & MIS Pro")
        st.caption("Dhan, Zerodha, FYERS, Groww — पोझिशन सायझिंग, ब्रेकआऊट व ट्रेंडलाईन चार्ट्स आणि ट्रेड जर्नल")
    with top_c2:
        st.write("")
        if st.button("🔄 रिफ्रेश करा", key="global_refresh_btn", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    df_stocks = load_stocks_and_margins()
    if df_stocks.empty:
        st.error("Nifty_500_Stocks.xlsx सापडली नाही! कृपया GitHub रेपॉझिटरी तपासा.")
        return

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "🧮 कॅल्क्युलेटर",
        "🚀 ब्रेकआऊट रिपोर्ट्स",
        "📈 ब्रेकआऊट चार्ट्स",
        "📐 ट्रेंडलाईन चार्ट्स",
        "📊 प्राईस-व्हॉल्यूम",
        "📋 ५०० शेअर्स मार्जिन",
        "🏆 ट्रेड जर्नल (P&L)"
    ])

    with tab1:
        render_tab1_calculator(df_stocks)
    render_tab2_and_tab3(tab2, tab6, df_stocks)
    render_chart_scanner_tab(tab3, "breakout", "📈 Breakouts Scanner & Live Charts (Full Market Scan)")
    render_chart_scanner_tab(tab4, "trendline", "📐 Auto Trendline Analysis & Charts")
    render_price_volume_tab(tab5, df_stocks)
    render_tab4_journal(tab7)


if __name__ == "__main__":
    render_app()
