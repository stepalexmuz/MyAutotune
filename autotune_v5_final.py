# ================================================================
#  AUTOTUNE v5 FINAL -- Single integrated file
#  Tandem t:slim X2 + Dexcom G7 + Apple Watch 11 + Oura Ring 4
#  + Sakharny Dnevnik XML + Meal Timer (iOS Shortcuts)
#  Google Colab
#
#  DATA SOURCES:
#    1. Tidepool Excel      -- CGM, Bolus, Basal, Carb Ratios, ISF
#    2. Medication CSV      -- manual boluses
#    3. Sakharny Dnevnik    -- GI, fat, protein per meal (XML export)
#    4. Apple Watch         -- workouts (CSV), meal timer (TXT)
#    5. Oura Ring 4 API     -- readiness, sleep, temperature
#
#  MANUAL STEPS BEFORE EACH RUN:
#    1. Tidepool    → Export → Google Drive  (TidepoolExport.xlsx)
#    2. Dnevnik     → Export XML → Google Drive  (dnevnik_arhiv.xml)
#    3. meal_log    → Shortcut "Sync to Drive" → 1 tap  (meal_log.csv)
#    4. workouts    → Shortcut "Sync to Drive" → 1 tap  (workouts.csv)
#    5. Oura        → automatic via API
#
#  KEY CONSTANTS (empirical, user-validated):
#    ISF_CR_RATIO = 2.86  (ISF = CR / 2.86)
#    G_PER_XE     = 12    (grams per XE, Sakharny Dnevnik standard)
#    Safety cap   = dynamic (+-10% to +-30% per analysis cycle)
# ================================================================

from google.colab import drive
drive.mount("/content/drive")

# Uncomment on first run:
# !pip install reportlab scipy requests

import os
import io
import json
import time
import xml.etree.ElementTree as ET
import requests
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time, date
from urllib.parse import urlencode, urlparse, parse_qs
from scipy import stats as scipy_stats
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import ipywidgets as widgets
from IPython.display import display, clear_output
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from PIL import Image as PILImage


# ================================================================
# CONSTANTS
# ================================================================

DEFAULT_SLEEP_TARGET_BG  = 6.5
DEFAULT_DAY_TARGET_BG    = 6.1
DIA_HOURS                = 5.0

# ISF = CR / 2.86 -- empirical, user-validated, do NOT change
# without re-testing on personal glucose data
ISF_CR_RATIO             = 2.86

# K1 = 12 / CR -- per Sakharny Dnevnik (12 g per XE)
# Change to 10 or 11 only if you use a different XE definition
G_PER_XE                 = 12

MAX_CHANGE_FACTOR        = 0.30
MAX_CARB_ABSORB_HOURS    = 4.0
COB_MIN_THRESHOLD        = 0.5
MIN_CSF_PER_GRAM         = 0.05

TIR_LOW_STD              = 3.9
TIR_HIGH_STD             = 10.0

MIN_SEG_CGM_PTS          = 5
MIN_BASAL_PTS            = 8
MIN_ISF_PTS              = 8
MIN_MEALS_FOR_CR         = 1

HIGH_CONF_BASAL_PTS      = 20
HIGH_CONF_ISF_PTS        = 20
HIGH_CONF_MEALS          = 5

BASAL_MEAN_BG_TOLERANCE  = 0.5
DEV_NOISE_FLOOR          = 0.05

MIN_DAYS_FOR_TREND       = 3
TREND_SLOPE_THRESHOLD    = 0.5
TREND_CV_HIGH            = 20.0

# Activity
ACTIVITY_BG_FACTOR       = 0.015  # mmol/L per kcal, auto-calibrated
POST_ACTIVITY_HOURS      = 6.0
SENSITIVITY_PEAK         = 1.25

# Oura
OURA_CLIENT_ID           = "9b7f81bb-009c-4754-854f-a0dff955cdfb"
TEMP_DEV_FLAG            = 0.3    # degrees C above personal baseline

# Readiness weights: 10 grades
READINESS_WEIGHTS = [
    (90, 1.00), (85, 0.95), (80, 0.90), (75, 0.85), (70, 0.80),
    (65, 0.73), (60, 0.66), (55, 0.59), (50, 0.50), (0,  0.38),
]

# Meal absorption thresholds (minutes from carb entry to CGM peak)
FAST_ABSORPTION_MIN      = 25
SLOW_ABSORPTION_MIN      = 90

# CSF analysis window per absorption type (hours)
CSF_WINDOWS = {
    "fast":    1.5,
    "normal":  3.0,
    "slow":    4.0,
    "unknown": 3.0,
}

# Dnevnik matching tolerance (minutes)
DNEVNIK_MATCH_MIN        = 45
WATCH_MATCH_MIN          = 45


# ================================================================
# FILE PATHS
# ================================================================

BASE_PATH        = "/content/drive/MyDrive/MyDiabet/ai_diabetes_project"
FILE_PATH        = BASE_PATH + "/TidepoolExport.xlsx"
MED_FILE_PATH    = BASE_PATH + "/medication_data_1.csv"
CONFIG_FILE_PATH = BASE_PATH + "/autotune_config.json"
WATCH_CSV_PATH   = BASE_PATH + "/workouts.csv"
MEAL_LOG_PATH    = BASE_PATH + "/meal_log.csv"
DNEVNIK_XML_PATH = BASE_PATH + "/dnevnik_arhiv.xml"
OURA_TOKEN_FILE  = BASE_PATH + "/oura_token.json"


# ================================================================
# CONFIGURATION
# ================================================================

def load_config():
    defaults = {
        "sleep_start":          "23:00",
        "sleep_end":            "07:00",
        "activity_bg_factor":   ACTIVITY_BG_FACTOR,
        "sensitivity_peak":     SENSITIVITY_PEAK,
        "oura_client_secret":   "",
    }
    if os.path.exists(CONFIG_FILE_PATH):
        try:
            with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            for k, v in defaults.items():
                cfg.setdefault(k, v)
            return cfg
        except Exception:
            pass
    return defaults

def save_config(cfg):
    with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ================================================================
# TIME UTILITIES
# ================================================================

def str_to_time(s):
    return datetime.strptime(s, "%H:%M").time()

def time_in_interval(t, start, end):
    if start <= end:
        return start <= t < end
    return t >= start or t < end

def is_sleep_segment(start_t, sleep_start, sleep_end):
    return time_in_interval(start_t, sleep_start, sleep_end)

def get_profile_value_at_time(profile_df, t, value_col, start_col="Start"):
    if profile_df.empty:
        return np.nan
    if start_col not in profile_df.columns or value_col not in profile_df.columns:
        return np.nan
    appl = profile_df[profile_df[start_col] <= t].sort_values(
        start_col, ascending=False)
    if appl.empty:
        appl = profile_df.sort_values(start_col)
    try:
        return float(appl.iloc[0][value_col])
    except Exception:
        return np.nan

def find_isf_value_col(pisf):
    for kw in ["insulin sensitivity amount", "sensitivity amount", "isf amount"]:
        for c in pisf.columns:
            if kw in c.lower():
                return c
    for c in pisf.columns:
        if "amount" in c.lower():
            return c
    skip = {"schedule name","start","insulin sensitivity start",
            "sensitivity start","device time"}
    for c in pisf.columns:
        if c.lower() not in skip:
            return c
    return None


# ================================================================
# OURA OAUTH2
# ================================================================

OURA_AUTH_URL  = "https://cloud.ouraring.com/oauth/authorize"
OURA_TOKEN_URL = "https://api.ouraring.com/oauth/token"
OURA_API_BASE  = "https://api.ouraring.com/v2/usercollection"
OURA_SCOPES    = "personal daily spo2"

def _load_oura_token():
    if os.path.exists(OURA_TOKEN_FILE):
        try:
            with open(OURA_TOKEN_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return None

def _save_oura_token(t):
    with open(OURA_TOKEN_FILE, "w") as f:
        json.dump(t, f, indent=2)

def _refresh_oura_token(token_data, client_secret):
    resp = requests.post(OURA_TOKEN_URL, data={
        "grant_type":    "refresh_token",
        "refresh_token": token_data["refresh_token"],
        "client_id":     OURA_CLIENT_ID,
        "client_secret": client_secret,
    })
    if resp.status_code == 200:
        t = resp.json()
        t["expires_at"] = time.time() + t.get("expires_in", 86400)
        _save_oura_token(t)
        return t
    raise RuntimeError("Oura refresh failed: " + resp.text)

def _get_oura_client_secret():
    """
    Read Oura Client Secret from Colab Secrets (recommended)
    or fall back to config file.
    Add secret in Colab: key=OURA_CLIENT_SECRET, value=<your secret>
    """
    try:
        from google.colab import userdata
        secret = userdata.get("OURA_CLIENT_SECRET")
        if secret:
            return secret.strip()
    except Exception:
        pass
    # Fallback: config file
    cfg = load_config()
    return cfg.get("oura_client_secret", "")

def get_oura_access_token(client_secret=None):
    """
    Returns valid Oura access token.
    client_secret: if None, reads from Colab Secrets automatically.

    First run requires OAuth2 browser authorization.
    Subsequent runs refresh automatically from stored token.

    HOW TO AUTHORIZE (first run only):
      1. Run the script -- authorization URL will be printed
      2. Open URL in browser, log in to Oura, click Authorize
      3. Browser redirects to http://localhost/?code=XXXX
         (page shows error -- that is OK)
      4. Copy the FULL URL from browser address bar
      5. Paste into the text widget that appears below the output
      6. Token is saved -- future runs are automatic
    """
    if client_secret is None:
        client_secret = _get_oura_client_secret()
    if not client_secret:
        raise ValueError(
            "Oura Client Secret not found.\n"
            "Add it in Colab: Secrets (key icon) -> "
            "Add secret: OURA_CLIENT_SECRET = <your secret>")

    token = _load_oura_token()
    if token:
        if time.time() < token.get("expires_at", 0) - 300:
            return token["access_token"]
        try:
            return _refresh_oura_token(token, client_secret)["access_token"]
        except Exception as e:
            print("Token refresh failed (" + str(e) + ") -- re-authorizing...")

    params = {"response_type":"code","client_id":OURA_CLIENT_ID,
              "redirect_uri":"http://localhost","scope":OURA_SCOPES,"state":"autotune"}
    auth_url = OURA_AUTH_URL + "?" + urlencode(params)

    print("=" * 65)
    print("OURA AUTHORIZATION REQUIRED  (one-time only)")
    print("=" * 65)
    print("1. Open this URL in your browser:")
    print("")
    print("   " + auth_url)
    print("")
    print("2. Authorize -> browser redirects to http://localhost/?code=...")
    print("3. Paste the full redirect URL into the box below and press Enter")
    print("=" * 65)

    # Use Colab ipywidgets text input -- works inside widget Output context
    url_input  = widgets.Text(
        placeholder="Paste full redirect URL here",
        layout=widgets.Layout(width="600px"))
    submit_btn = widgets.Button(description="Submit", button_style="primary")
    auth_out   = widgets.Output()
    result_box = {"code": None}

    def on_submit(b):
        with auth_out:
            url = url_input.value.strip()
            code = parse_qs(urlparse(url).query).get("code", [None])[0]
            if code:
                result_box["code"] = code
                print("Code received. Exchanging for token...")
            else:
                print("ERROR: could not find code= in the URL. Try again.")

    submit_btn.on_click(on_submit)
    display(widgets.HBox([url_input, submit_btn]))
    display(auth_out)

    # Wait up to 120 seconds for user to paste URL
    import time as _time
    for _ in range(240):
        _time.sleep(0.5)
        if result_box["code"]:
            break
    else:
        raise TimeoutError("Oura authorization timed out (120s). Run again.")

    code = result_box["code"]
    resp = requests.post(OURA_TOKEN_URL, data={
        "grant_type":"authorization_code","code":code,
        "redirect_uri":"http://localhost",
        "client_id":OURA_CLIENT_ID,"client_secret":client_secret,
    })
    if resp.status_code != 200:
        raise RuntimeError("Token exchange failed: " + resp.text)
    t = resp.json()
    t["expires_at"] = time.time() + t.get("expires_in", 86400)
    _save_oura_token(t)
    print("Oura authorization successful. Token saved.")
    return t["access_token"]

def oura_get(endpoint, access_token, start_date, end_date):
    headers = {"Authorization": "Bearer " + access_token}
    params  = {"start_date": str(start_date), "end_date": str(end_date)}
    resp = requests.get(OURA_API_BASE + "/" + endpoint,
                        headers=headers, params=params, timeout=30)
    if resp.status_code == 200:
        return resp.json().get("data", [])
    print("Oura API warning: " + endpoint + " HTTP " + str(resp.status_code))
    return []


# ================================================================
# OURA DATA LOADING AND PREPROCESSING
# ================================================================

def load_oura_data(access_token, start_dt, end_dt):
    s = start_dt.date() - timedelta(days=1)
    e = end_dt.date()
    raw_r = oura_get("daily_readiness", access_token, s, e)
    raw_s = oura_get("daily_sleep",     access_token, s, e)

    def to_df(records, cols):
        if not records:
            return pd.DataFrame(columns=["date"] + cols)
        rows = []
        for r in records:
            row = {"date": pd.to_datetime(
                r.get("day", r.get("date",""))).date()}
            for c in cols:
                val = r.get(c)
                if val is None:
                    val = r.get("contributors",{}).get(c)
                row[c] = val
            rows.append(row)
        return pd.DataFrame(rows)

    return {
        "readiness": to_df(raw_r, ["score","temperature_deviation",
                                    "temperature_trend_deviation"]),
        "sleep":     to_df(raw_s, ["score","efficiency",
                                    "total_sleep_duration"]),
    }

def get_readiness_weight(score):
    if pd.isna(score):
        return 0.75
    score = float(score)
    for threshold, weight in READINESS_WEIGHTS:
        if score >= threshold:
            return weight
    return 0.38

def preprocess_oura(oura_data, start_dt, end_dt):
    rd = oura_data.get("readiness", pd.DataFrame())
    sl = oura_data.get("sleep",     pd.DataFrame())
    result = {}
    current = start_dt.date()
    while current <= end_dt.date():
        entry = {
            "readiness_weight":      0.80,
            "readiness_score":       np.nan,
            "temperature_deviation": np.nan,
            "sleep_efficiency":      np.nan,
            "possibly_stressed":     False,
            "temp_flagged":          False,
        }
        if not rd.empty and "date" in rd.columns:
            row = rd[rd["date"] == current]
            if not row.empty:
                score    = row.iloc[0].get("score", np.nan)
                temp_dev = row.iloc[0].get("temperature_deviation", np.nan)
                entry["readiness_score"]       = score
                entry["readiness_weight"]      = get_readiness_weight(score)
                entry["temperature_deviation"] = temp_dev
                entry["possibly_stressed"]     = (
                    pd.notna(score) and float(score) < 50)
                entry["temp_flagged"]          = (
                    pd.notna(temp_dev) and
                    abs(float(temp_dev)) > TEMP_DEV_FLAG)
        if not sl.empty and "date" in sl.columns:
            row = sl[sl["date"] == current]
            if not row.empty:
                entry["sleep_efficiency"] = row.iloc[0].get("efficiency", np.nan)
        result[current] = entry
        current += timedelta(days=1)
    return result


# ================================================================
# SAKHARNY DNEVNIK XML PARSER
# ================================================================

def compute_absorption_score(gi, fat_g, protein_g, carbs_g):
    if pd.isna(gi) or gi <= 0:
        return np.nan
    fat_g     = float(fat_g)     if pd.notna(fat_g)     else 0.0
    protein_g = float(protein_g) if pd.notna(protein_g) else 0.0
    carbs_g   = float(carbs_g)   if pd.notna(carbs_g) and carbs_g > 0 else 1.0
    fat_f  = min(0.30, (fat_g     / carbs_g) * 0.5)
    prot_f = min(0.20, (protein_g / carbs_g) * 0.3)
    return float(gi) * (1.0 - fat_f) * (1.0 - prot_f)

def classify_absorption(score):
    if pd.isna(score):
        return "unknown"
    if score > FAST_ABSORPTION_MIN * 2.6:   # > 65
        return "fast"
    if score < SLOW_ABSORPTION_MIN * 0.44:  # < 40
        return "slow"
    return "normal"

def get_csf_window(absorption_type):
    return CSF_WINDOWS.get(absorption_type, CSF_WINDOWS["unknown"])

def parse_dnevnik_xml(xml_path, start_dt, end_dt):
    if not os.path.exists(xml_path):
        print("INFO: Dnevnik XML not found at " + xml_path)
        return pd.DataFrame()

    t_min = start_dt - timedelta(hours=1)
    t_max = end_dt   + timedelta(hours=1)
    records = []

    try:
        context = ET.iterparse(xml_path, events=("start",))
        for event, elem in context:
            if elem.tag != "RECORD":
                elem.clear()
                continue
            if elem.get("idrec") is not None:
                elem.clear()
                continue
            date_str = elem.get("date","")
            time_str = elem.get("time","")
            if not date_str or not time_str:
                elem.clear()
                continue
            try:
                ts = datetime.strptime(date_str + " " + time_str, "%d.%m.%Y %H:%M")
            except ValueError:
                elem.clear()
                continue
            if ts < t_min or ts > t_max:
                elem.clear()
                continue

            def _f(attr, default=np.nan):
                val = elem.get(attr,"")
                try:
                    return float(val) if val not in ("","0",None) else default
                except (ValueError, TypeError):
                    return default

            gi  = _f("gi")
            fat = _f("zh", 0.0)
            prt = _f("b",  0.0)
            u   = _f("u",  0.0)
            score = compute_absorption_score(gi, fat, prt, u)

            records.append({
                "timestamp":        ts,
                "carbs_g":          u,
                "gi":               gi,
                "gn":               _f("gn"),
                "fat_g":            fat,
                "protein_g":        prt,
                "kcal":             _f("kkal", 0.0),
                "k1_used":          _f("k1"),
                "ki_calculated":    _f("ki"),
                "he":               _f("he"),
                "absorption_score": score,
                "absorption_type":  classify_absorption(score),
            })
            elem.clear()
    except ET.ParseError as e:
        print("Dnevnik XML parse error: " + str(e))
        return pd.DataFrame()

    if not records:
        print("INFO: No Dnevnik records in selected period")
        return pd.DataFrame()

    df = pd.DataFrame(records).sort_values("timestamp").reset_index(drop=True)
    print("Dnevnik: " + str(len(df)) + " meal records loaded")
    return df


# ================================================================
# APPLE WATCH MEAL LOG PARSER
# ================================================================

def parse_meal_log(log_path, start_dt, end_dt):
    if not os.path.exists(log_path):
        print("INFO: meal_log.csv not found at " + log_path)
        return []
    try:
        # No header in file
        df = pd.read_csv(log_path, header=None, names=["datetime","event"])
    except Exception as e:
        print("meal_log read error: " + str(e))
        return []

    df["datetime"] = pd.to_datetime(df["datetime"], errors="coerce")
    df = (df.dropna(subset=["datetime"])
            .sort_values("datetime")
            .reset_index(drop=True))
    df = df[
        (df["datetime"] >= start_dt - timedelta(hours=1)) &
        (df["datetime"] <= end_dt   + timedelta(hours=1))
    ].copy()

    if df.empty:
        return []

    pairs = []
    pending_start = None

    for _, row in df.iterrows():
        evt = str(row["event"]).strip().lower()
        ts  = row["datetime"]
        if evt == "start":
            if pending_start is not None:
                pairs.append({"meal_start": pending_start,
                              "meal_stop": pd.NaT,
                              "meal_duration_min": np.nan})
            pending_start = ts
        elif evt == "stop":
            if pending_start is not None:
                dur = (ts - pending_start).total_seconds() / 60.0
                pairs.append({"meal_start": pending_start,
                              "meal_stop": ts,
                              "meal_duration_min": round(dur, 1)})
                pending_start = None

    if pending_start is not None:
        pairs.append({"meal_start": pending_start,
                      "meal_stop": pd.NaT,
                      "meal_duration_min": np.nan})

    print("Watch meal log: " + str(len(pairs)) + " session(s) parsed")
    return pairs


# ================================================================
# MEAL ENRICHMENT
# ================================================================

def enrich_meals(meals_df, dnevnik_df, watch_pairs):
    if meals_df.empty:
        return meals_df

    df = meals_df.copy()
    df["Timestamp"] = pd.to_datetime(df["Timestamp"])

    for col in ["gi","gn","fat_g","protein_g","kcal",
                "absorption_score","absorption_type",
                "meal_start","meal_stop","meal_duration_min",
                "bolus_delta_min","reliability",
                "dnevnik_matched","watch_matched"]:
        if col not in df.columns:
            df[col] = (np.nan if col not in
                       ("absorption_type","reliability") else "unknown")
    df["dnevnik_matched"] = False
    df["watch_matched"]   = False
    df["reliability"]     = "unknown"

    for idx, row in df.iterrows():
        bolus_ts = pd.Timestamp(row["Timestamp"])

        # Dnevnik match
        if not dnevnik_df.empty and "timestamp" in dnevnik_df.columns:
            deltas = (dnevnik_df["timestamp"] - bolus_ts).dt.total_seconds().abs() / 60.0
            best_i = deltas.idxmin()
            if deltas[best_i] <= DNEVNIK_MATCH_MIN:
                b = dnevnik_df.loc[best_i]
                df.at[idx, "gi"]               = b.get("gi",       np.nan)
                df.at[idx, "gn"]               = b.get("gn",       np.nan)
                df.at[idx, "fat_g"]            = b.get("fat_g",    np.nan)
                df.at[idx, "protein_g"]        = b.get("protein_g",np.nan)
                df.at[idx, "kcal"]             = b.get("kcal",     np.nan)
                df.at[idx, "absorption_score"] = b.get("absorption_score", np.nan)
                df.at[idx, "absorption_type"]  = b.get("absorption_type",  "unknown")
                df.at[idx, "dnevnik_matched"]  = True
                df.at[idx, "reliability"]      = "measured"

        # Watch match -- watch_pairs is a DataFrame
        if not (watch_pairs is None or
                (hasattr(watch_pairs, "empty") and watch_pairs.empty)):
            best_pair  = None
            best_delta = WATCH_MATCH_MIN + 1
            for _, pair in watch_pairs.iterrows():
                if pd.isna(pair.get("meal_start")):
                    continue
                ms    = pd.Timestamp(pair["meal_start"])
                delta = abs((bolus_ts - ms).total_seconds()) / 60.0
                if delta < best_delta:
                    best_delta = delta
                    best_pair  = pair
            if best_pair is not None and best_delta <= WATCH_MATCH_MIN:
                ms = pd.Timestamp(best_pair["meal_start"])
                df.at[idx, "meal_start"]        = ms
                df.at[idx, "meal_stop"]         = best_pair.get("meal_stop", pd.NaT)
                df.at[idx, "meal_duration_min"] = best_pair.get("duration_min", np.nan)
                df.at[idx, "bolus_delta_min"]   = round(
                    (bolus_ts - ms).total_seconds() / 60.0, 1)
                df.at[idx, "watch_matched"]     = True

    n_d = int(df["dnevnik_matched"].sum())
    n_w = int(df["watch_matched"].sum())
    print("Meal enrichment: " + str(len(df)) + " meals | Dnevnik: " +
          str(n_d) + " | Watch: " + str(n_w))
    return df

def get_segment_absorption_type(enriched_meals_df, st, et):
    if enriched_meals_df.empty or "absorption_type" not in enriched_meals_df.columns:
        return "unknown"
    seg = enriched_meals_df[
        enriched_meals_df["Timestamp"].dt.time.apply(
            lambda t: time_in_interval(t, st, et))
    ].copy()
    if seg.empty:
        return "unknown"
    measured = seg[seg["reliability"] == "measured"] if "reliability" in seg.columns else seg
    source   = measured if not measured.empty else seg
    counts   = source["absorption_type"].value_counts()
    if counts.empty:
        return "unknown"
    top = counts.index[0]
    if top == "unknown" and len(counts) > 1:
        top = counts.index[1]
    return top

def analyze_bolus_timing(enriched_meals_df):
    if enriched_meals_df.empty:
        return {}
    df = enriched_meals_df[enriched_meals_df["watch_matched"] == True].copy()
    if df.empty:
        return {"n_with_watch": 0}
    deltas = df["bolus_delta_min"].dropna()
    if deltas.empty:
        return {"n_with_watch": len(df)}
    by_abs = {}
    for atype, grp in df.groupby("absorption_type"):
        d = grp["bolus_delta_min"].dropna()
        if not d.empty:
            by_abs[atype] = round(float(d.mean()), 1)
    return {
        "n_with_watch":   len(df),
        "mean_delta_min": round(float(deltas.mean()), 1),
        "pre_bolus_pct":  round(100.0 * float((deltas < 0).mean()), 1),
        "late_bolus_pct": round(100.0 * float((deltas > 10).mean()), 1),
        "by_absorption":  by_abs,
    }


# ================================================================
# MEAL LOG EDITOR (run as separate cell when needed)
# ================================================================

def open_meal_log_editor():
    print("""
# ================================================================
#  MEAL LOG EDITOR -- copy and run in a new cell
# ================================================================
import pandas as pd
from IPython.display import display

MEAL_LOG_PATH = "/content/drive/MyDrive/MyDiabet/ai_diabetes_project/meal_log.csv"

df = pd.read_csv(MEAL_LOG_PATH, header=None, names=["datetime","event"])

# To add a row:
# new = pd.DataFrame([{"datetime":"2026-05-02 12:30:00","event":"start"}])
# df = pd.concat([df, new]).sort_values("datetime").reset_index(drop=True)

# To delete a row by index:
# df = df.drop(index=5).reset_index(drop=True)

# To save:
# df.to_csv(MEAL_LOG_PATH, header=False, index=False)
# print("Saved.")
""")


# ================================================================
# APPLE WATCH WORKOUTS
# ================================================================

def load_watch_workouts(csv_path, start_dt, end_dt):
    if not os.path.exists(csv_path):
        print("INFO: workouts.csv not found -- no activity data")
        return pd.DataFrame(columns=["type","start","end","calories","duration_min"])
    try:
        df = pd.read_csv(csv_path)
    except Exception as e:
        print("workouts.csv read error: " + str(e))
        return pd.DataFrame(columns=["type","start","end","calories","duration_min"])

    df.columns = [c.lower().strip() for c in df.columns]
    for col in ["start","end"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if "start" in df.columns:
        df = df[(df["start"] >= start_dt) & (df["start"] <= end_dt)]
    if "calories" in df.columns:
        df["calories"] = pd.to_numeric(df["calories"], errors="coerce").fillna(0)
    else:
        df["calories"] = 0.0
    if "duration_min" not in df.columns:
        if "duration" in df.columns:
            df["duration_min"] = pd.to_numeric(df["duration"], errors="coerce").fillna(0)
        elif "start" in df.columns and "end" in df.columns:
            df["duration_min"] = ((df["end"] - df["start"])
                                  .dt.total_seconds() / 60.0).fillna(0)
        else:
            df["duration_min"] = 0.0
    return df.reset_index(drop=True)

def build_activity_windows(workouts_df):
    windows = []
    if workouts_df.empty:
        return windows
    for _, row in workouts_df.iterrows():
        start = pd.Timestamp(row["start"])
        end   = (pd.Timestamp(row["end"]) if pd.notna(row.get("end"))
                 else start + pd.Timedelta(minutes=float(row.get("duration_min",30))))
        cal   = float(row.get("calories", 0))
        dur_m = max(1.0, (end - start).total_seconds() / 60.0)
        windows.append({
            "type":         str(row.get("type","Workout")),
            "start":        start,
            "end":          end,
            "cal_per_5min": cal / (dur_m / 5.0),
            "post_end":     end + pd.Timedelta(hours=POST_ACTIVITY_HOURS),
        })
    return windows

def compute_activity_bgi(cgm_df, activity_windows, isf_arr, bg_factor):
    n    = len(cgm_df)
    abgi = np.zeros(n)
    if not activity_windows:
        return abgi
    times = cgm_df["Device Time"].values
    for i, t in enumerate(times):
        ts = pd.Timestamp(t)
        for w in activity_windows:
            if w["start"] <= ts <= w["end"]:
                abgi[i] -= w["cal_per_5min"] * bg_factor
                break
    return abgi

def compute_isf_multiplier(cgm_df, activity_windows):
    n   = len(cgm_df)
    mul = np.ones(n)
    if not activity_windows:
        return mul
    times = cgm_df["Device Time"].values
    for i, t in enumerate(times):
        ts   = pd.Timestamp(t)
        best = 1.0
        for w in activity_windows:
            if w["start"] <= ts <= w["end"]:
                best = max(best, SENSITIVITY_PEAK)
            elif w["end"] < ts <= w["post_end"]:
                hours_after = (ts - w["end"]).total_seconds() / 3600.0
                m = 1.0 + (SENSITIVITY_PEAK - 1.0) * (1.0 - hours_after / POST_ACTIVITY_HOURS)
                best = max(best, m)
        mul[i] = best
    return mul

def build_activity_mask(cgm_df, activity_windows):
    n          = len(cgm_df)
    in_workout = np.zeros(n, dtype=bool)
    in_post    = np.zeros(n, dtype=bool)
    if not activity_windows:
        return in_workout, in_post
    times = cgm_df["Device Time"].values
    for i, t in enumerate(times):
        ts = pd.Timestamp(t)
        for w in activity_windows:
            if w["start"] <= ts <= w["end"]:
                in_workout[i] = True
                break
            if w["end"] < ts <= w["post_end"]:
                in_post[i] = True
                break
    return in_workout, in_post

def calibrate_activity_factor(cgm_df, activity_windows):
    if not activity_windows or cgm_df.empty:
        return ACTIVITY_BG_FACTOR
    factors = []
    for w in activity_windows:
        pts = cgm_df[
            (cgm_df["Device Time"] >= w["start"]) &
            (cgm_df["Device Time"] <= w["end"]) &
            (cgm_df["deviation"].notna())
        ]
        if len(pts) < 3:
            continue
        mean_res = float(pts["deviation"].mean())
        if w["cal_per_5min"] > 0:
            factors.append(-mean_res / w["cal_per_5min"])
    if len(factors) >= 3:
        cal = float(np.median(factors))
        cal = max(0.005, min(0.08, cal))
        print("Activity calibration: factor = " + str(round(cal, 4)) +
              " (" + str(len(factors)) + " workouts)")
        return cal
    print("Activity calibration: insufficient data, using default " +
          str(ACTIVITY_BG_FACTOR))
    return ACTIVITY_BG_FACTOR


# ================================================================
# TIDEPOOL + MEDICATION DATA LOADING
# ================================================================

def diagnose_tidepool_fixed(fp):
    """Повна діагностика + безпечний парсинг"""
    if not os.path.exists(fp):
        print("❌ Файл відсутній")
        return
    
    xls = pd.ExcelFile(fp)
    print(f"✅ {len(xls.sheet_names)} листів:", xls.sheet_names)
    
    data = {}
    for sheet in xls.sheet_names:
        if sheet == 'EXPORT ERROR': continue
        try:
            df = xls.parse(sheet, nrows=3)
            print(f"\n{sheet}:")
            print(f"  Рядків: {len(df)}, Колонок: {len(df.columns)}")
            if len(df) > 0 and len(df.columns) > 0:
                print(f"  Колонки: {list(df.columns)}")
                print(f"  Приклад: {dict(df.iloc[0].dropna())}")
            data[sheet] = df
        except Exception as e:
            print(f"  ❌ {e}")
    
    return data

# Запуск
tidepool_data = diagnose_tidepool_fixed(FILE_PATH)

def load_tidepool_data(fp):
    xls = pd.ExcelFile(fp)
    def _parse(sheet, col="Device Time"):
        if sheet not in xls.sheet_names:
            return pd.DataFrame()
        df = xls.parse(sheet)
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
        return df
    return (
        _parse("Basal Schedules"),
        _parse("Bolus"),
        _parse("CGM"),
        _parse("Carb Ratios"),
        _parse("Insulin Sensitivities"),
        _parse("Basal"),
    )

def load_tidepool_data_fixed(fp):
    """Виправлена версія для t:slim X2"""
    xls = pd.ExcelFile(fp)
    
    def _parse_safe(sheet, col="Device Time"):
        if sheet not in xls.sheet_names or sheet == 'EXPORT ERROR':
            return pd.DataFrame()
        try:
            df = xls.parse(sheet)
            time_cols = ["Device Time", "normalTime", "displayTime", "Timestamp"]
            for tc in time_cols:
                if tc in df.columns:
                    df["Device Time"] = pd.to_datetime(df[tc], errors="coerce")
                    break
            return df
        except:
            return pd.DataFrame()
    
    return (
        _parse_safe("Basal Schedules"),
        _parse_safe("Bolus"),
        _parse_safe("CGM"), 
        _parse_safe("Carb Ratios"),
        _parse_safe("Insulin Sensitivities"),
        _parse_safe("Basal"),
    )

# ТЕСТ
basal_sch, bolus_df, cgm_df, carb_ratios, isf_df, basal_df = load_tidepool_data_fixed(FILE_PATH)
print("✅ FIXED Tidepool:")
print(f"  CGM: {len(cgm_df)} points")
print(f"  Bolus: {len(bolus_df)}")
print(f"  Basal: {len(basal_df)}")
print("  CGM колонки:", list(cgm_df.columns)[:10])


def load_bolus_calculator(fp):
    xls = pd.ExcelFile(fp)
    if "Bolus Calculator" not in xls.sheet_names:
        return pd.DataFrame(columns=["Timestamp","Carbs"])
    df = xls.parse("Bolus Calculator")
    ts_col = next((c for c in ["Device Time","Timestamp","Local Time"]
                   if c in df.columns), None)
    if not ts_col:
        return pd.DataFrame(columns=["Timestamp","Carbs"])
    df["Timestamp"] = pd.to_datetime(df[ts_col], errors="coerce")
    carb_col = None
    for c in df.columns:
        if "carb" in c.lower() and "input" in c.lower():
            carb_col = c; break
    if not carb_col:
        for c in df.columns:
            if "carb" in c.lower():
                carb_col = c; break
    if not carb_col:
        return pd.DataFrame(columns=["Timestamp","Carbs"])
    df["Carbs"] = pd.to_numeric(df[carb_col], errors="coerce")
    return (df[["Timestamp","Carbs"]].dropna()
              .query("Carbs > 0").reset_index(drop=True))

def load_medication_data(fp):
    with open(fp, "r", encoding="utf-8") as f:
        lines = f.readlines()
    data_start = 0
    for i, line in enumerate(lines):
        if "Timestamp,Name,Value,Medication Type" in line:
            data_start = i; break
    df = pd.read_csv(fp, skiprows=data_start, header=0)
    if df.columns[0] != "Timestamp":
        df = df.rename(columns={df.columns[0]: "Timestamp"})
    df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
    df["Dose"]      = pd.to_numeric(df["Value"], errors="coerce")
    return df[["Timestamp","Dose"]].dropna()


# ================================================================
# BOLUS AGGREGATION
# ================================================================

def prepare_bolus_minute(bolus_manual_df, bolus_pump_df):
    entries = []
    if bolus_pump_df is not None and not bolus_pump_df.empty:
        for _, row in bolus_pump_df.iterrows():
            dt  = pd.to_datetime(row.get("Device Time"), errors="coerce")
            if pd.isna(dt): continue
            sub  = str(row.get("Sub Type","normal")).lower()
            norm = float(row.get("Normal",   0) or 0)
            ext  = float(row.get("Extended", 0) or 0)
            dur  = float(row.get("Duration (mins)", 0) or 0)
            if norm > 0:
                entries.append({"Timestamp":dt,"Dose":norm,"Type":sub})
            if sub in ("extended","dual/square") and ext > 0 and dur > 0:
                dpm = round(ext / dur, 5)
                for m in range(int(dur)):
                    entries.append({"Timestamp":dt+timedelta(minutes=m),
                                    "Dose":dpm,"Type":sub})
    if bolus_manual_df is not None and not bolus_manual_df.empty:
        for _, row in bolus_manual_df.iterrows():
            ts   = pd.to_datetime(row.get("Timestamp"), errors="coerce")
            dose = round(float(row.get("Value", row.get("Dose",0)) or 0), 5)
            if not pd.isna(ts) and dose > 0:
                entries.append({"Timestamp":ts,"Dose":dose,"Type":"manual"})
    if not entries:
        return pd.DataFrame(columns=["Timestamp","Dose","Type"])
    df = (pd.DataFrame(entries).sort_values("Timestamp")
            .dropna(subset=["Timestamp"]).reset_index(drop=True))
    df["Timestamp"] = pd.to_datetime(df["Timestamp"])
    return df


# ================================================================
# TDD
# ================================================================

def calculate_tdd_avg(bolus_min_df, basal_deliv_df, start_dt, end_dt):
    n_days = max(1, (end_dt - start_dt).days)
    basal_total = 0.0
    if not basal_deliv_df.empty:
        bd = basal_deliv_df.copy()
        tc = next((c for c in ["Timestamp","Device Time","Local Time","Zulu Time"]
                   if c in bd.columns), None)
        if tc:
            bd["_ts"] = pd.to_datetime(bd[tc], errors="coerce")
            bd = bd.dropna(subset=["_ts"])
            bd = bd[(bd["_ts"] >= start_dt) & (bd["_ts"] <= end_dt)]
            rc = next((c for c in ["Rate","Value","Basal","Basal Rate","Delivered"]
                       if c in bd.columns), None)
            if rc and "Duration (mins)" in bd.columns:
                basal_total = float(
                    (pd.to_numeric(bd[rc], errors="coerce") *
                     pd.to_numeric(bd["Duration (mins)"], errors="coerce") / 60.0
                     ).sum())
    bolus_total = 0.0
    if not bolus_min_df.empty:
        mask = ((bolus_min_df["Timestamp"] >= start_dt) &
                (bolus_min_df["Timestamp"] <= end_dt))
        bolus_total = float(bolus_min_df.loc[mask,"Dose"].sum())
    tdd = (basal_total + bolus_total) / n_days
    return {
        "tdd_avg":     round(tdd, 2),
        "basal_total": round(basal_total, 2),
        "bolus_total": round(bolus_total, 2),
        "n_days":      n_days,
        "cr_base":     round(500.0/tdd, 2) if tdd > 0 else np.nan,
        "isf_base":    round(100.0/tdd, 2) if tdd > 0 else np.nan,
    }


# ================================================================
# BGI ENGINE
# ================================================================

def calculate_bgi_vectorized(bolus_min_df, cgm_times_series,
                              isf_arr, isf_multiplier_arr,
                              dia_hours=DIA_HOURS):
    dia_min = dia_hours * 60.0
    times   = pd.to_datetime(cgm_times_series)
    n       = len(times)
    bgi     = np.zeros(n)
    if bolus_min_df.empty or n == 0:
        return bgi
    cgm_arr = times.values.astype("datetime64[s]").astype(np.float64)
    eff_isf = isf_arr * isf_multiplier_arr
    for _, row in bolus_min_df.iterrows():
        dose = float(row["Dose"])
        if dose <= 0: continue
        try:
            bt = np.datetime64(pd.Timestamp(row["Timestamp"]),"s").astype(np.float64)
        except Exception:
            continue
        mins = (cgm_arr - bt) / 60.0
        rate = np.zeros(n)
        rate[(mins > 30) & (mins <= 75)]      = 1.0 / 45.0
        rate[(mins > 75) & (mins <= dia_min)] = 1.0 / 225.0
        bgi -= dose * rate * eff_isf * 5.0
    return bgi

def calculate_basal_bgi(basal_rate_arr, isf_arr, isf_multiplier_arr):
    return -(np.asarray(basal_rate_arr, float) *
             np.asarray(isf_arr, float) *
             np.asarray(isf_multiplier_arr, float) * (5.0/60.0))

def calculate_delta_bg(cgm_df):
    times   = cgm_df["Device Time"].values.astype("datetime64[s]").astype(np.float64)
    values  = cgm_df["Value"].values.astype(float)
    n       = len(times)
    delta   = np.full(n, np.nan)
    window_s  = 15 * 60
    max_gap_s = 25 * 60
    for i in range(n):
        t_now  = times[i]
        t_back = t_now - window_s
        cands  = np.where(times <= t_back)[0]
        if len(cands) == 0: continue
        j    = cands[-1]
        dt_s = t_now - times[j]
        if dt_s <= 0 or dt_s > max_gap_s: continue
        delta[i] = (values[i] - values[j]) / (dt_s / 300.0)
    return delta


# ================================================================
# COB + MEAL EXCLUSION + ATTRIBUTION
# ================================================================

def compute_cob_series(cgm_df, meals_df):
    n   = len(cgm_df)
    cob = np.zeros(n)
    if meals_df.empty:
        return cob
    for _, meal in meals_df.iterrows():
        mt    = pd.Timestamp(meal["Timestamp"])
        carbs = float(meal["Carbs"])
        if carbs <= 0: continue
        window_end = mt + pd.Timedelta(hours=MAX_CARB_ABSORB_HOURS)
        mask = ((cgm_df["Device Time"] >= mt) &
                (cgm_df["Device Time"] <= window_end)).values
        idxs = np.where(mask)[0]
        if len(idxs) == 0: continue
        min_step = carbs / len(idxs)
        cob_rem  = carbs
        for idx in idxs:
            if cob_rem <= COB_MIN_THRESHOLD: break
            cob[idx] += cob_rem
            dev    = cgm_df.iloc[idx].get("deviation",  np.nan)
            cr_val = cgm_df.iloc[idx].get("cr_profile", np.nan)
            isf_v  = cgm_df.iloc[idx].get("isf_profile",np.nan)
            extra  = 0.0
            if (pd.notna(dev) and dev > 0 and pd.notna(cr_val) and
                    cr_val > 0 and pd.notna(isf_v) and isf_v > 0):
                extra = dev * cr_val / isf_v
            cob_rem = max(0.0, cob_rem - min(cob_rem, min_step + extra))
    return cob

def build_meal_exclusion_mask(cgm_df, meals_df):
    n    = len(cgm_df)
    excl = np.zeros(n, dtype=bool)
    if meals_df.empty:
        return excl
    for _, meal in meals_df.iterrows():
        mt         = pd.Timestamp(meal["Timestamp"])
        window_end = mt + pd.Timedelta(hours=MAX_CARB_ABSORB_HOURS)
        excl |= ((cgm_df["Device Time"] >= mt) &
                 (cgm_df["Device Time"] <= window_end)).values
    return excl

def attribute_cgm_points(cgm_df):
    n    = len(cgm_df)
    attr = np.array(["Basal"] * n, dtype=object)
    cols = set(cgm_df.columns)
    for i in range(n):
        dev  = cgm_df.iloc[i]["deviation"]  if "deviation"  in cols else np.nan
        bgi  = cgm_df.iloc[i]["bgi"]        if "bgi"        in cols else np.nan
        bbgi = cgm_df.iloc[i]["basal_bgi"]  if "basal_bgi"  in cols else np.nan
        delt = cgm_df.iloc[i]["delta_bg"]   if "delta_bg"   in cols else np.nan
        cob  = cgm_df.iloc[i]["cob"]        if "cob"        in cols else 0.0
        if pd.isna(dev) or pd.isna(bgi):
            attr[i] = "Basal"; continue
        if cob > COB_MIN_THRESHOLD and dev > 0:
            attr[i] = "CSF"; continue
        if bgi >= 0:
            attr[i] = "Basal"; continue
        if pd.notna(bbgi) and bbgi < 0 and abs(bgi) <= abs(bbgi)/4.0:
            attr[i] = "Basal"; continue
        if pd.notna(delt) and delt > 0:
            attr[i] = "Basal"; continue
        attr[i] = "ISF"
    return attr


# ================================================================
# PROBLEM DETECTION
# ================================================================

def detect_bolus_overlaps(bolus_pump_df, bolus_manual_df, window_min=15):
    overlaps = []
    if bolus_pump_df.empty or bolus_manual_df.empty: return overlaps
    if "Sub Type" not in bolus_pump_df.columns: return overlaps
    auto = bolus_pump_df[
        bolus_pump_df["Sub Type"].str.contains("automated",case=False,na=False)
    ].copy()
    if auto.empty: return overlaps
    auto["_t"] = pd.to_datetime(auto["Device Time"], errors="coerce")
    auto = auto.dropna(subset=["_t"])
    manual = bolus_manual_df.copy()
    manual["_t"] = pd.to_datetime(manual["Timestamp"], errors="coerce")
    manual = manual.dropna(subset=["_t"])
    for _, r in auto.iterrows():
        t0 = r["_t"]; t1 = t0 + timedelta(minutes=window_min)
        if not manual[(manual["_t"]>=t0)&(manual["_t"]<=t1)].empty:
            overlaps.append((t0,t1))
    return overlaps

def detect_basal_stops(bd_df, threshold_min=30):
    stops = []
    if bd_df.empty or "Rate" not in bd_df.columns: return stops
    tc = next((c for c in ["Device Time","Timestamp","Local Time","Zulu Time"]
               if c in bd_df.columns), None)
    if not tc: return stops
    bd = bd_df.copy()
    bd["_t"] = pd.to_datetime(bd[tc], errors="coerce")
    bd = bd.dropna(subset=["_t"]).sort_values("_t")
    z  = bd[bd["Rate"]==0].copy()
    if z.empty: return stops
    z["_d"] = z["_t"].diff() > timedelta(minutes=5)
    z["_g"] = z["_d"].cumsum()
    for _, grp in z.groupby("_g"):
        s, e = grp["_t"].min(), grp["_t"].max()
        dur  = (e-s).total_seconds()/60.0
        if dur >= threshold_min:
            stops.append((s,e,dur))
    return stops


# ================================================================
# HELPERS
# ================================================================

def k1_from_cr(cr):
    return round(float(G_PER_XE)/float(cr),3) if (cr and pd.notna(cr) and cr>0) else np.nan

def apply_cap(rec_full, profile_val, cap_factor):
    if pd.isna(rec_full) or pd.isna(profile_val) or profile_val==0 or cap_factor==0:
        return np.nan, False
    delta     = rec_full - profile_val
    max_delta = abs(profile_val) * cap_factor
    if abs(delta) > max_delta:
        return profile_val + np.sign(delta)*max_delta, True
    return rec_full, False

def median_iqr_filter(arr):
    if len(arr) < 4:
        return arr, float(np.median(arr)), 0
    q25, q75 = np.percentile(arr,25), np.percentile(arr,75)
    iqr  = q75 - q25
    mask = (arr >= q25-1.5*iqr) & (arr <= q75+1.5*iqr)
    filtered  = arr[mask]
    n_removed = int(np.sum(~mask))
    med = float(np.median(filtered)) if len(filtered)>0 else float(np.median(arr))
    return filtered, med, n_removed

def weighted_median(values, weights):
    values  = np.asarray(values,  dtype=float)
    weights = np.asarray(weights, dtype=float)
    valid   = ~np.isnan(values) & (weights > 0)
    if valid.sum() == 0: return np.nan
    values  = values[valid]; weights = weights[valid]
    order   = np.argsort(values)
    values  = values[order]; weights = weights[order]
    cum     = np.cumsum(weights)
    return float(values[cum >= cum[-1]/2.0][0])

def _fmt(v, d=2):
    return str(round(float(v),d)) if pd.notna(v) else "?"

def confidence_label(n_basal, n_isf, n_meals):
    if (n_basal>=HIGH_CONF_BASAL_PTS and n_isf>=HIGH_CONF_ISF_PTS
            and n_meals>=HIGH_CONF_MEALS):
        return "High"
    if n_basal>=MIN_BASAL_PTS and n_isf>=MIN_ISF_PTS:
        return "Medium"
    if n_basal>=MIN_BASAL_PTS or n_isf>=MIN_ISF_PTS:
        return "Low"
    return "Insufficient"

def detect_cr_isf_contradiction(d_cr, d_isf):
    if pd.isna(d_cr) or pd.isna(d_isf): return False
    if abs(d_cr)<0.01 or abs(d_isf)<0.001: return False
    return (d_cr > 0) != (d_isf > 0)

def dynamic_cap(confidence, basal_cv, trend_sig, contradiction):
    if confidence == "Insufficient": return 0.0
    if confidence == "High":
        cap = 0.20 if (pd.notna(basal_cv) and basal_cv>=TREND_CV_HIGH) else 0.30
    elif confidence == "Medium":
        cap = 0.20
    else:
        cap = 0.10
    if contradiction:
        cap = min(cap, 0.05 if confidence=="Low" else 0.10)
    if trend_sig:
        cap = max(0.05, cap-0.05)
    return cap

def analyze_basal_trend(basal_deliv_df, start_dt, end_dt):
    base = {"daily_df":pd.DataFrame(columns=["date","basal_U"]),
            "slope":np.nan,"cv_pct":np.nan,"trend_label":"insufficient",
            "trend_sig":False,"cv_high":False,"n_days":0}
    if basal_deliv_df.empty: return base
    bd = basal_deliv_df.copy()
    tc = next((c for c in ["Timestamp","Device Time","Local Time","Zulu Time"]
               if c in bd.columns), None)
    if not tc: return base
    bd["_ts"] = pd.to_datetime(bd[tc], errors="coerce")
    bd = bd.dropna(subset=["_ts"])
    bd = bd[(bd["_ts"]>=start_dt)&(bd["_ts"]<=end_dt)]
    rc = next((c for c in ["Rate","Value","Basal","Basal Rate","Delivered"]
               if c in bd.columns), None)
    if not rc or "Duration (mins)" not in bd.columns: return base
    bd["_u"] = (pd.to_numeric(bd[rc],errors="coerce") *
                pd.to_numeric(bd["Duration (mins)"],errors="coerce") / 60.0)
    bd["_date"] = bd["_ts"].dt.date
    daily = bd.groupby("_date")["_u"].sum().reset_index()
    daily.columns = ["date","basal_U"]
    daily = daily.sort_values("date").reset_index(drop=True)
    n = len(daily)
    if n < MIN_DAYS_FOR_TREND:
        base["daily_df"]=daily; base["n_days"]=n; return base
    x = np.arange(n, dtype=float)
    y = daily["basal_U"].values.astype(float)
    slope, *_ = scipy_stats.linregress(x, y)
    cv_pct = float(np.std(y)/np.mean(y)*100.0) if np.mean(y)>0 else np.nan
    trend_sig = abs(slope) > TREND_SLOPE_THRESHOLD
    return {"daily_df":daily,"slope":round(slope,3),
            "cv_pct":round(cv_pct,1) if pd.notna(cv_pct) else np.nan,
            "trend_label":"stable" if not trend_sig else ("increasing" if slope>0 else "decreasing"),
            "trend_sig":trend_sig,
            "cv_high":pd.notna(cv_pct) and cv_pct>TREND_CV_HIGH,
            "n_days":n}


# ================================================================
# MAIN AUTOTUNE ENGINE
# ================================================================

def run_autotune(profile_name, start_dt, end_dt,
                 sleep_start_str, sleep_end_str,
                 oura_context, activity_windows,
                 meals_enriched):

    sleep_start = str_to_time(sleep_start_str)
    sleep_end   = str_to_time(sleep_end_str)

    basal_sched, bolus_df, cgm_raw, cr_df, isf_df, basal_deliv = \
        load_tidepool_data(FILE_PATH)
    med_df   = load_medication_data(MED_FILE_PATH)
    meals_df = load_bolus_calculator(FILE_PATH)
    meals_df = meals_df[
        (meals_df["Timestamp"] >= start_dt) &
        (meals_df["Timestamp"] <= end_dt)
    ].reset_index(drop=True)

    basal_trend = analyze_basal_trend(basal_deliv, start_dt, end_dt)

    pb = basal_sched[basal_sched["Schedule Name"]==profile_name].copy()
    if pb.empty: raise ValueError("Profile not found: " + profile_name)
    pb = (pb.drop_duplicates("Basal Schedule Start",keep="first")
            .sort_values("Basal Schedule Start").reset_index(drop=True))
    if len(pb) > 16: pb = pb.iloc[:16].reset_index(drop=True)
    pb["StartTime"] = pd.to_datetime(
        pb["Basal Schedule Start"],errors="coerce").dt.time

    pcr = cr_df[cr_df["Schedule Name"]==profile_name].copy().reset_index(drop=True)
    if "Carb Ratio Start" in pcr.columns:
        pcr["Start"] = pd.to_datetime(pcr["Carb Ratio Start"],errors="coerce").dt.time

    pisf = isf_df[isf_df["Schedule Name"]==profile_name].copy().reset_index(drop=True)
    for col_nm in ["Insulin Sensitivity Start","Sensitivity Start"]:
        if col_nm in pisf.columns:
            pisf["Start"] = pd.to_datetime(pisf[col_nm],errors="coerce").dt.time; break
    isf_val_col = find_isf_value_col(pisf)

    cgm_raw = cgm_raw.copy()
    cgm_raw["Device Time"] = pd.to_datetime(cgm_raw["Device Time"],errors="coerce")
    cgm_raw = (cgm_raw.dropna(subset=["Device Time"])
                      .sort_values("Device Time").reset_index(drop=True))
    cgm = cgm_raw[
        (cgm_raw["Device Time"]>=start_dt) &
        (cgm_raw["Device Time"]<=end_dt)
    ].copy().reset_index(drop=True)

    if cgm.empty: raise ValueError("No CGM data in selected period")
    if "Value" not in cgm.columns: raise ValueError("CGM has no Value column")
    cgm["time_only"] = cgm["Device Time"].dt.time

    bolus_min = prepare_bolus_minute(med_df, bolus_df)
    bolus_for_bgi = (bolus_min[
        (bolus_min["Timestamp"] >= start_dt-timedelta(hours=DIA_HOURS)) &
        (bolus_min["Timestamp"] <= end_dt)
    ].copy() if not bolus_min.empty else bolus_min.copy())

    def cr_at(t_obj):
        return get_profile_value_at_time(pcr, t_obj, "Carb Ratio Amount")
    def isf_at(t_obj):
        return get_profile_value_at_time(pisf, t_obj, isf_val_col) if isf_val_col else np.nan
    def basal_at(t_obj):
        return get_profile_value_at_time(pb, t_obj, "Basal Schedule Rate",
                                          start_col="StartTime")

    cgm["cr_profile"]    = cgm["time_only"].apply(cr_at)
    cgm["isf_profile"]   = cgm["time_only"].apply(isf_at)
    cgm["basal_profile"] = cgm["time_only"].apply(basal_at)

    isf_arr   = cgm["isf_profile"].values.astype(float)
    basal_arr = cgm["basal_profile"].values.astype(float)

    tdd_info = calculate_tdd_avg(bolus_min, basal_deliv, start_dt, end_dt)

    # Activity
    act_bg_factor = load_config().get("activity_bg_factor", ACTIVITY_BG_FACTOR)
    isf_mul = compute_isf_multiplier(cgm, activity_windows)
    in_workout, in_post_act = build_activity_mask(cgm, activity_windows)
    cgm["isf_multiplier"] = isf_mul
    cgm["in_workout"]     = in_workout
    cgm["in_post_act"]    = in_post_act

    # Readiness weight
    def get_rw(dt):
        d   = pd.Timestamp(dt).date()
        ctx = oura_context.get(d, {})
        return ctx.get("readiness_weight", 0.80)
    cgm["readiness_weight"] = cgm["Device Time"].apply(get_rw)

    # BGI + deviation
    cgm["bgi"]       = calculate_bgi_vectorized(
        bolus_for_bgi, cgm["Device Time"], isf_arr, isf_mul)
    cgm["basal_bgi"] = calculate_basal_bgi(basal_arr, isf_arr, isf_mul)
    cgm["abgi"]      = compute_activity_bgi(
        cgm, activity_windows, isf_arr, act_bg_factor)
    cgm["delta_bg"]  = calculate_delta_bg(cgm)
    cgm["deviation"] = cgm["delta_bg"] - cgm["bgi"] - cgm["abgi"]

    # Calibrate activity factor
    cal_factor = calibrate_activity_factor(cgm, activity_windows)
    if cal_factor != act_bg_factor:
        cgm["abgi"]      = compute_activity_bgi(cgm, activity_windows, isf_arr, cal_factor)
        cgm["deviation"] = cgm["delta_bg"] - cgm["bgi"] - cgm["abgi"]
        cfg = load_config(); cfg["activity_bg_factor"] = cal_factor; save_config(cfg)

    cgm["cob"]           = compute_cob_series(cgm, meals_df)
    cgm["attribution"]   = attribute_cgm_points(cgm)
    cgm["meal_excluded"] = build_meal_exclusion_mask(cgm, meals_df)

    overlaps    = detect_bolus_overlaps(bolus_df, med_df)
    basal_stops = detect_basal_stops(basal_deliv)

    start_times = pb["StartTime"].tolist()
    n_segs      = len(start_times)

    # Pre-compute CR recs
    cr_recs_full = {}
    cr_n_meals   = {}
    cr_sources   = {}

    for i in range(n_segs):
        st = start_times[i]
        if st is None:
            cr_recs_full[i]=np.nan; cr_sources[i]="no segment"; cr_n_meals[i]=0; continue
        et = start_times[i+1] if i+1<n_segs else time(23,59,59)
        cr_pv  = cr_at(st); isf_pv = isf_at(st)
        meals_seg = meals_df[
            meals_df["Timestamp"].dt.time.apply(lambda t: time_in_interval(t,st,et))
        ].copy()
        abs_type  = get_segment_absorption_type(meals_enriched, st, et)
        csf_hours = get_csf_window(abs_type)

        if len(meals_seg)<MIN_MEALS_FOR_CR or pd.isna(cr_pv) or pd.isna(isf_pv):
            cr_recs_full[i]=np.nan; cr_sources[i]="no meals"; cr_n_meals[i]=0; continue

        meal_cr_recs = []
        for _, meal in meals_seg.iterrows():
            mt    = meal["Timestamp"]
            carbs = float(meal["Carbs"])
            csf_pts = cgm[
                (cgm["attribution"]=="CSF") &
                (cgm["Device Time"]>=mt) &
                (cgm["Device Time"]<=mt+pd.Timedelta(hours=csf_hours)) &
                (cgm["deviation"].notna())
            ]
            if csf_pts.empty or carbs<=0: continue
            actual_csf = float(csf_pts["deviation"].sum()) / carbs
            if actual_csf >= MIN_CSF_PER_GRAM:
                meal_cr_recs.append(isf_pv / actual_csf)

        if meal_cr_recs:
            cr_recs_full[i] = float(np.mean(meal_cr_recs))
            cr_sources[i]   = "calculated (" + abs_type + ")"
            cr_n_meals[i]   = len(meal_cr_recs)
        else:
            cr_recs_full[i]=np.nan; cr_sources[i]="no CSF signal"; cr_n_meals[i]=0

    calc_idxs = [i for i in range(n_segs) if pd.notna(cr_recs_full.get(i))]
    for i in range(n_segs):
        if pd.notna(cr_recs_full.get(i)): continue
        st = start_times[i]
        if st is None: continue
        prev_list = [j for j in calc_idxs if j<i]
        next_list = [j for j in calc_idxs if j>i]
        if prev_list and next_list:
            pi,ni = prev_list[-1],next_list[0]
            pt_m = start_times[pi].hour*60+start_times[pi].minute
            nt_m = start_times[ni].hour*60+start_times[ni].minute
            st_m = st.hour*60+st.minute
            if nt_m != pt_m:
                frac = (st_m-pt_m)/(nt_m-pt_m)
                cr_recs_full[i] = cr_recs_full[pi]+frac*(cr_recs_full[ni]-cr_recs_full[pi])
            else:
                cr_recs_full[i] = cr_recs_full[pi]
            cr_sources[i]="interpolated"; cr_n_meals[i]=0
        elif prev_list:
            cr_recs_full[i]=cr_recs_full[prev_list[-1]]; cr_sources[i]="extrapolated"; cr_n_meals[i]=0
        elif next_list:
            cr_recs_full[i]=cr_recs_full[next_list[0]]; cr_sources[i]="extrapolated"; cr_n_meals[i]=0

    # Main loop
    tech_rows = []

    for i in range(n_segs):
        st = start_times[i]
        if st is None: continue
        et = start_times[i+1] if i+1<n_segs else time(23,59,59)
        interval_str = st.strftime("%H:%M")+"-"+et.strftime("%H:%M")

        cr_pv    = cr_at(st)
        isf_pv   = isf_at(st)
        basal_pv = (float(pb.iloc[i]["Basal Schedule Rate"])
                    if "Basal Schedule Rate" in pb.columns else np.nan)

        seg = cgm[cgm["time_only"].apply(lambda t: time_in_interval(t,st,et))].copy()
        target_bg = (DEFAULT_SLEEP_TARGET_BG
                     if is_sleep_segment(st,sleep_start,sleep_end)
                     else DEFAULT_DAY_TARGET_BG)

        if len(seg) < MIN_SEG_CGM_PTS:
            tech_rows.append(_empty_tech_row(
                interval_str, target_bg, basal_pv, cr_pv, isf_pv, tdd_info,
                cr_recs_full.get(i), cr_sources.get(i), "not enough CGM data"))
            continue

        vals    = seg["Value"].values.astype(float)
        vals_ok = vals[~np.isnan(vals)]
        mean_bg  = float(np.nanmean(vals_ok)) if len(vals_ok)>0 else np.nan
        std_bg   = float(np.nanstd(vals_ok))  if len(vals_ok)>1 else np.nan
        cv_pct   = (std_bg/mean_bg*100.0 if pd.notna(std_bg) and mean_bg>0 else np.nan)
        tir_pers = (100.0*float(np.mean((vals_ok>=target_bg-1.5)&(vals_ok<=target_bg+1.5)))
                    if len(vals_ok)>0 else np.nan)
        tir_std  = (100.0*float(np.mean((vals_ok>=TIR_LOW_STD)&(vals_ok<=TIR_HIGH_STD)))
                    if len(vals_ok)>0 else np.nan)

        n_csf = int((seg["attribution"]=="CSF").sum())
        n_act = int((seg["in_workout"]|seg["in_post_act"]).sum())

        # Basal pts (meal-excluded)
        basal_pts = cgm[
            (cgm["attribution"]=="Basal") &
            (~cgm["meal_excluded"]) &
            (cgm["time_only"].apply(lambda t: time_in_interval(t,st,et))) &
            (cgm["deviation"].notna())
        ]
        n_basal_pts = len(basal_pts)
        basal_full = basal_safe = np.nan
        basal_dev_median = np.nan
        basal_capped = basal_implausible = False
        basal_in_tolerance = basal_noise_floor = False

        if n_basal_pts>=MIN_BASAL_PTS and pd.notna(basal_pv):
            devs    = basal_pts["deviation"].values.astype(float)
            devs_ok = devs[~np.isnan(devs)]
            w_arr   = basal_pts["readiness_weight"].values[~np.isnan(devs)]
            if len(devs_ok) >= MIN_BASAL_PTS:
                basal_dev_median = weighted_median(devs_ok, w_arr)
                bg_vals = basal_pts["Value"].values.astype(float)
                bg_vals = bg_vals[~np.isnan(bg_vals)]
                basal_mean_bg = float(np.nanmean(bg_vals)) if len(bg_vals)>0 else np.nan
                basal_noise_floor  = abs(basal_dev_median) < DEV_NOISE_FLOOR
                basal_in_tolerance = (pd.notna(basal_mean_bg) and
                                      abs(basal_mean_bg-target_bg)<=BASAL_MEAN_BG_TOLERANCE)
                if not basal_noise_floor and not basal_in_tolerance:
                    avg_isf = float(basal_pts["isf_profile"].median())
                    if avg_isf > 0:
                        basal_offset   = basal_dev_median*12.0/avg_isf
                        basal_full_raw = basal_pv + basal_offset
                        if abs(basal_full_raw-basal_pv) > 2.0*abs(basal_pv):
                            basal_implausible = True
                            basal_full = basal_pv+np.sign(basal_full_raw-basal_pv)*2.0*abs(basal_pv)
                        else:
                            basal_full = basal_full_raw

        # ISF pts (meal-excluded)
        isf_pts = cgm[
            (cgm["attribution"]=="ISF") &
            (~cgm["meal_excluded"]) &
            (cgm["time_only"].apply(lambda t: time_in_interval(t,st,et))) &
            (cgm["deviation"].notna())
        ]
        n_isf_pts = len(isf_pts)
        isf_full = isf_safe = np.nan
        isf_capped = False

        if n_isf_pts>=MIN_ISF_PTS and pd.notna(isf_pv) and isf_pv>0:
            isf_devs    = isf_pts["deviation"].values.astype(float)
            isf_devs_ok = isf_devs[~np.isnan(isf_devs)]
            w_isf       = isf_pts["readiness_weight"].values[~np.isnan(isf_devs)]
            if len(isf_devs_ok) >= MIN_ISF_PTS:
                isf_dev_median = weighted_median(isf_devs_ok, w_isf)
                if abs(isf_dev_median) >= DEV_NOISE_FLOOR:
                    denom = isf_pv + isf_dev_median
                    isf_full = (isf_pv*isf_pv)/denom if abs(denom)>0.01 else isf_pv

        # CR
        cr_full_val = cr_recs_full.get(i, np.nan)
        n_meals     = cr_n_meals.get(i, 0)
        cr_src      = cr_sources.get(i, "no data")
        d_cr_full   = (cr_full_val-cr_pv) if pd.notna(cr_full_val) and pd.notna(cr_pv) else np.nan
        d_isf_full  = (isf_full-isf_pv)   if pd.notna(isf_full)   and pd.notna(isf_pv) else np.nan

        conf         = confidence_label(n_basal_pts, n_isf_pts, n_meals)
        contradiction = detect_cr_isf_contradiction(d_cr_full, d_isf_full)
        cap          = dynamic_cap(conf, basal_trend["cv_pct"] or 0.0,
                                   basal_trend["trend_sig"], contradiction)

        basal_safe, basal_capped = (apply_cap(basal_full, basal_pv, cap)
                                    if pd.notna(basal_full) else (np.nan,False))
        isf_safe,   isf_capped   = (apply_cap(isf_full,   isf_pv,   cap)
                                    if pd.notna(isf_full)  else (np.nan,False))
        cr_safe,    cr_capped    = (apply_cap(cr_full_val, cr_pv,    cap)
                                    if pd.notna(cr_full_val) else (np.nan,False))

        d_basal = (basal_safe-basal_pv) if pd.notna(basal_safe) and pd.notna(basal_pv) else np.nan
        d_isf   = (isf_safe-isf_pv)     if pd.notna(isf_safe)   and pd.notna(isf_pv)   else np.nan
        d_cr    = (cr_safe-cr_pv)       if pd.notna(cr_safe)     and pd.notna(cr_pv)     else np.nan

        if pd.notna(mean_bg) and mean_bg>target_bg+1.5: status="hyper"
        elif pd.notna(mean_bg) and mean_bg<target_bg-1.5: status="hypo"
        elif pd.notna(tir_pers) and tir_pers>=70 and (pd.isna(cv_pct) or cv_pct<=36):
            status="optimal"
        else: status="suboptimal"

        abs_type = get_segment_absorption_type(meals_enriched, st, et)

        def _r(v, d=3):
            return round(float(v),d) if pd.notna(v) else np.nan

        tech_rows.append({
            "Interval":           interval_str,
            "Target BG":          round(target_bg,1),
            "Mean BG":            _r(mean_bg,2),
            "SD":                 _r(std_bg,2),
            "CV %":               _r(cv_pct,1),
            "TIR personal %":     _r(tir_pers,1),
            "TIR std %":          _r(tir_std,1),
            "N Basal":            n_basal_pts,
            "N ISF":              n_isf_pts,
            "N CSF":              n_csf,
            "N Meals":            n_meals,
            "N Activity pts":     n_act,
            "Absorption":         abs_type,
            "Basal Profile":      _r(basal_pv),
            "Basal Rec Safe":     _r(basal_safe),
            "Basal Rec Full":     _r(basal_full),
            "D Basal":            _r(d_basal),
            "CR Profile":         _r(cr_pv,2),
            "CR Rec Safe":        _r(cr_safe,2),
            "CR Rec Full":        _r(cr_full_val,2),
            "D CR":               _r(d_cr,2),
            "CR Source":          cr_src,
            "ISF Profile":        _r(isf_pv),
            "ISF Rec Safe":       _r(isf_safe),
            "ISF Rec Full":       _r(isf_full),
            "D ISF":              _r(d_isf),
            "K1 Profile":         k1_from_cr(cr_pv),
            "K1 Rec Safe":        k1_from_cr(cr_safe),
            "Cap Applied %":      round(cap*100),
            "Basal Capped":       basal_capped,
            "CR Capped":          cr_capped,
            "ISF Capped":         isf_capped,
            "Basal Implausible":  basal_implausible,
            "Basal In Tolerance": basal_in_tolerance,
            "Basal Noise Floor":  basal_noise_floor,
            "Contradiction":      contradiction,
            "Confidence":         conf,
            "Status":             status,
            "Overlap":  any(ps.hour==st.hour for ps,_ in overlaps),
            "BasalStop":any(ss.hour==st.hour for ss,_,__ in basal_stops),
        })

    tech_df = pd.DataFrame(tech_rows)

    all_vals = cgm["Value"].values.astype(float)
    all_vals = all_vals[~np.isnan(all_vals)]
    overall = {
        "mean_bg":             round(float(np.nanmean(all_vals)),2) if len(all_vals)>0 else np.nan,
        "tir_std":             round(100.0*float(np.mean(
                                (all_vals>=TIR_LOW_STD)&(all_vals<=TIR_HIGH_STD))),1)
                               if len(all_vals)>0 else np.nan,
        "cv":                  round(float(np.nanstd(all_vals)/np.nanmean(all_vals)*100),1)
                               if len(all_vals)>1 else np.nan,
        "n_workouts":          len(activity_windows),
        "n_activity_pts":      int(in_workout.sum()+in_post_act.sum()),
        "calibrated_bg_factor":round(cal_factor,4),
    }

    return tech_df, cgm, tdd_info, overall, basal_trend, cal_factor

def _empty_tech_row(interval_str, target_bg, basal_pv, cr_pv, isf_pv,
                    tdd_info, cr_full=None, cr_source="no data", note=""):
    def _r(v,d=3):
        return round(float(v),d) if (v is not None and pd.notna(v)) else np.nan
    return {
        "Interval":interval_str,"Target BG":target_bg,
        "Mean BG":np.nan,"SD":np.nan,"CV %":np.nan,
        "TIR personal %":np.nan,"TIR std %":np.nan,
        "N Basal":0,"N ISF":0,"N CSF":0,"N Meals":0,
        "N Activity pts":0,"Absorption":"unknown",
        "Basal Profile":_r(basal_pv),"Basal Rec Safe":np.nan,
        "Basal Rec Full":np.nan,"D Basal":np.nan,
        "CR Profile":_r(cr_pv,2),"CR Rec Safe":np.nan,
        "CR Rec Full":_r(cr_full,2) if cr_full is not None else np.nan,
        "D CR":np.nan,"CR Source":cr_source,
        "ISF Profile":_r(isf_pv),"ISF Rec Safe":np.nan,
        "ISF Rec Full":np.nan,"D ISF":np.nan,
        "K1 Profile":k1_from_cr(cr_pv),"K1 Rec Safe":np.nan,
        "Cap Applied %":0,
        "Basal Capped":False,"CR Capped":False,"ISF Capped":False,
        "Basal Implausible":False,"Basal In Tolerance":False,
        "Basal Noise Floor":False,"Contradiction":False,
        "Confidence":"Insufficient","Status":"no data",
        "Overlap":False,"BasalStop":False,
    }


# ================================================================
# ACTION TABLE + SUMMARY
# ================================================================

def build_action_table(tech_df):
    rows = []
    for _, r in tech_df.iterrows():
        conf = r["Confidence"]
        if conf == "Insufficient": continue
        has_basal = pd.notna(r["D Basal"]) and abs(r["D Basal"])>0.001
        has_cr    = pd.notna(r["D CR"])    and abs(r["D CR"])>0.01
        has_isf   = pd.notna(r["D ISF"])   and abs(r["D ISF"])>0.001
        changes = []
        if has_basal:
            arrow = "UP" if r["D Basal"]>0 else "DOWN"
            extra = (" (REVIEW)" if r["Basal Implausible"] else
                     " (tolerance)" if r["Basal In Tolerance"] else
                     " (noise)" if r["Basal Noise Floor"] else
                     " (capped)" if r["Basal Capped"] else "")
            changes.append("Basal "+arrow+" "+_fmt(abs(r["D Basal"]),3)+" U/h"+extra)
        if has_cr:
            arrow = "UP" if r["D CR"]>0 else "DOWN"
            note  = (" ["+r["CR Source"]+"]"
                     if r["CR Source"] in ("interpolated","extrapolated") else "")
            changes.append("CR "+arrow+" "+_fmt(abs(r["D CR"]),2)+" g/U"+
                           (" (capped)" if r["CR Capped"] else "")+note)
        if has_isf:
            arrow = "UP" if r["D ISF"]>0 else "DOWN"
            changes.append("ISF "+arrow+" "+_fmt(abs(r["D ISF"]),3)+
                           (" (capped)" if r["ISF Capped"] else ""))
        action = " | ".join(changes) if changes else "No change needed"
        flags = []
        if r.get("Contradiction"):
            flags.append("CR/ISF CONTRADICTION cap="+str(r.get("Cap Applied %","?"))+"%")
        if r["Overlap"]:   flags.append("bolus-overlap")
        if r["BasalStop"]: flags.append("basal-stop")
        if flags:
            action = "WARNING: "+" | ".join(flags)+" || "+action
        rows.append({
            "Segment":    r["Interval"],
            "Status":     r["Status"].upper(),
            "Mean BG":    r["Mean BG"],
            "TIR %":      r["TIR personal %"],
            "CV %":       r["CV %"],
            "Cap %":      r.get("Cap Applied %","-"),
            "Absorption": r.get("Absorption","-"),
            "Act pts":    r.get("N Activity pts",0),
            "Basal now":  r["Basal Profile"],
            "Basal ->":   r["Basal Rec Safe"] if has_basal else "-",
            "CR now":     r["CR Profile"],
            "CR ->":      r["CR Rec Safe"]    if has_cr    else "-",
            "ISF now":    r["ISF Profile"],
            "ISF ->":     r["ISF Rec Safe"]   if has_isf   else "-",
            "K1 now":     r["K1 Profile"],
            "K1 ->":      r["K1 Rec Safe"]    if has_cr    else "-",
            "Confidence": conf,
            "Action":     action,
        })
    return pd.DataFrame(rows)


def build_summary(tech_df, tdd_info, overall, basal_trend,
                  oura_context, meals_enriched, timing_summary,
                  profile_name, start_dt, end_dt):
    lines = []
    lines.append("AUTOTUNE v5 FINAL -- ANALYSIS SUMMARY")
    lines.append("Profile: "+profile_name+"   |   "+
                 str(start_dt.date())+" to "+str(end_dt.date())+
                 "   ("+str(tdd_info["n_days"])+" days)")
    lines.append("")

    lines.append("OVERALL GLUCOSE CONTROL")
    lines.append("  Average glucose : "+_fmt(overall["mean_bg"])+" mmol/L")
    lines.append("  Time In Range   : "+_fmt(overall["tir_std"])+
                 "% (standard 3.9-10.0 mmol/L)")
    lines.append("  Variability     : CV = "+_fmt(overall["cv"])+"%"+
                 ("  (acceptable)" if overall["cv"]<=36 else "  (HIGH)"))
    lines.append("  Total daily dose: "+str(tdd_info["tdd_avg"])+" U/day"+
                 "  (basal "+str(tdd_info["basal_total"])+
                 " U + bolus "+str(tdd_info["bolus_total"])+" U)")
    lines.append("  TDD reference:  CR="+str(tdd_info["cr_base"])+
                 " g/U  |  ISF="+str(tdd_info["isf_base"])+" mmol/U")
    lines.append("")

    lines.append("ACTIVITY  (Apple Watch)")
    lines.append("  Workouts: "+str(overall["n_workouts"])+
                 "  |  CGM pts in activity windows: "+str(overall["n_activity_pts"]))
    lines.append("  Calibrated activity factor: "+
                 str(overall["calibrated_bg_factor"])+" mmol/L per kcal")
    lines.append("")

    lines.append("OURA CONTEXT")
    if oura_context:
        scores  = [v["readiness_score"] for v in oura_context.values()
                   if pd.notna(v.get("readiness_score"))]
        temp_fl = [str(d) for d,v in oura_context.items() if v.get("temp_flagged")]
        stressed= [str(d) for d,v in oura_context.items() if v.get("possibly_stressed")]
        if scores:
            lines.append("  Avg readiness: "+_fmt(float(np.mean(scores)),0))
        if temp_fl:
            lines.append("  Temperature flagged: "+", ".join(temp_fl))
        if stressed:
            lines.append("  Low readiness (<50): "+", ".join(stressed))
        if not temp_fl and not stressed:
            lines.append("  No stress/temperature flags.")
    else:
        lines.append("  No Oura data.")
    lines.append("")

    lines.append("MEAL DATA")
    if not meals_enriched.empty:
        n_total = len(meals_enriched)
        n_d = int(meals_enriched["dnevnik_matched"].sum()) if "dnevnik_matched" in meals_enriched.columns else 0
        n_w = int(meals_enriched["watch_matched"].sum())   if "watch_matched"   in meals_enriched.columns else 0
        lines.append("  Total meals: "+str(n_total)+
                     "  |  Dnevnik GI: "+str(n_d)+
                     "  |  Watch timing: "+str(n_w))
        if "absorption_type" in meals_enriched.columns:
            counts = meals_enriched["absorption_type"].value_counts().to_dict()
            for t in ["fast","normal","slow","unknown"]:
                c = counts.get(t,0)
                if c > 0:
                    lines.append("    "+t.upper()+": "+str(c)+" meals")
        if timing_summary.get("n_with_watch",0) > 0:
            lines.append("  Pre-bolus timing (n="+str(timing_summary["n_with_watch"])+"):")
            lines.append("    Mean delta: "+str(timing_summary["mean_delta_min"])+
                         " min  (+ = bolus after meal start)")
            lines.append("    Pre-bolus: "+str(timing_summary["pre_bolus_pct"])+
                         "%  |  Late (>10 min): "+str(timing_summary["late_bolus_pct"])+"%")
    else:
        lines.append("  No meal data.")
    lines.append("")

    lines.append("BASAL TREND  ("+str(basal_trend["n_days"])+" days)")
    if basal_trend["trend_label"]=="insufficient":
        lines.append("  Insufficient data")
    else:
        lines.append("  Trend: "+basal_trend["trend_label"].upper()+
                     "  slope="+_fmt(basal_trend["slope"],2)+" U/day"+
                     ("  *** significant" if basal_trend["trend_sig"] else ""))
        lines.append("  CV: "+_fmt(basal_trend["cv_pct"],1)+"%"+
                     ("  (HIGH)" if basal_trend["cv_high"] else ""))
        if basal_trend["trend_label"]=="decreasing":
            lines.append("  >>> Decreasing need -- HIGHER PRIORITY for correction.")
    lines.append("")

    lines.append("SEGMENT FINDINGS")
    for label, col, val in [
        ("HIGH glucose",       "Status",           "hyper"),
        ("LOW glucose",        "Status",           "hypo"),
        ("Well-controlled",    "Status",           "optimal"),
        ("Insufficient data",  "Confidence",       "Insufficient"),
        ("Basal implausible",  "Basal Implausible", True),
        ("CR/ISF contradiction","Contradiction",    True),
    ]:
        if col not in tech_df.columns: continue
        subset = tech_df[tech_df[col]==val]
        if not subset.empty:
            lines.append("  "+label+": "+", ".join(subset["Interval"].tolist()))

    lines.append("")
    lines.append("OVERALL DIRECTION")
    has_ch = tech_df[tech_df["Confidence"].isin(["High","Medium","Low"])]
    for param, col, thr in [("Basal","D Basal",0.001),("CR","D CR",0.01),("ISF","D ISF",0.001)]:
        if col not in has_ch.columns: continue
        up   = has_ch[has_ch[col].fillna(0) >  thr]
        down = has_ch[has_ch[col].fillna(0) < -thr]
        line = "  "+param+" : "
        if not up.empty:   line += str(len(up))+" segment(s) UP"
        if not down.empty: line += ("  |  " if not up.empty else "")+str(len(down))+" segment(s) DOWN"
        if up.empty and down.empty: line += "no significant change"
        lines.append(line)

    lines.append("")
    lines.append("NOTES")
    lines.append("  - Safe rec: within dynamic cap (see Cap % column).")
    lines.append("  - Full rec: mathematical optimum (extended mode).")
    lines.append("  - Meal windows ("+str(MAX_CARB_ABSORB_HOURS)+
                 "h) excluded from basal/ISF analysis.")
    lines.append("  - Activity windows ("+str(POST_ACTIVITY_HOURS)+
                 "h post-workout) use modified ISF in BGI.")
    lines.append("  - Readiness weights applied to deviation medians.")
    lines.append("  - Verify all changes before programming the pump.")

    return "\n".join(lines)


# ================================================================
# CHARTS
# ================================================================

def build_summary_chart(cgm_aug, tech_df, tdd_info, oura_context,
                         activity_windows, profile_name, start_dt, end_dt):
    fig = plt.figure(figsize=(17,9))
    gs  = gridspec.GridSpec(3,1,height_ratios=[3.5,1.2,1.5],hspace=0.40)
    ax0 = fig.add_subplot(gs[0])
    ax1 = fig.add_subplot(gs[1])
    ax2 = fig.add_subplot(gs[2])

    if not cgm_aug.empty and "Value" in cgm_aug.columns:
        ax0.plot(cgm_aug["Device Time"],cgm_aug["Value"],
                 color="steelblue",lw=0.9,alpha=0.85,label="CGM (mmol/L)",zorder=5)

    ax0.axhspan(DEFAULT_DAY_TARGET_BG-1.5,DEFAULT_DAY_TARGET_BG+1.5,
                alpha=0.10,color="green",label="Personal target (day)")
    ax0.axhline(DEFAULT_SLEEP_TARGET_BG,color="purple",lw=0.7,ls="--",
                alpha=0.5,label="Night target")
    ax0.axhline(TIR_LOW_STD, color="orange",lw=0.6,ls=":",alpha=0.4)
    ax0.axhline(TIR_HIGH_STD,color="orange",lw=0.6,ls=":",alpha=0.4,
                label="ADA bounds")

    wl_added = pl_added = False
    for w in activity_windows:
        ax0.axvspan(w["start"],w["end"],alpha=0.20,color="dodgerblue",
                    label="Workout" if not wl_added else None,zorder=2)
        ax0.axvspan(w["end"],w["post_end"],alpha=0.08,color="dodgerblue",
                    label="Post-activity (+6h)" if not pl_added else None,zorder=1)
        try:
            ylim = ax0.get_ylim()
            ybot = ylim[0] if ylim[0]>0 else 3.5
        except Exception:
            ybot = 3.5
        ax0.annotate("▲ "+w["type"][:3],xy=(w["start"],ybot),
                     fontsize=6,color="dodgerblue",rotation=90,va="bottom")
        wl_added = pl_added = True

    if oura_context:
        current = start_dt.date()
        while current <= end_dt.date():
            ctx   = oura_context.get(current,{})
            score = ctx.get("readiness_score",75)
            if pd.notna(score):
                score = float(score)
                c = ("#2ecc71" if score>=75 else "#a8d08d" if score>=65 else
                     "#f9e04b" if score>=55 else "#f0a500" if score>=45 else "#e74c3c")
                d0 = pd.Timestamp(datetime.combine(current,datetime.min.time()))
                d1 = d0+pd.Timedelta(days=1)
                ax0.axvspan(d0,d1,ymin=0,ymax=0.04,alpha=0.7,color=c,zorder=6)
            if ctx.get("temp_flagged"):
                d_mid = pd.Timestamp(datetime.combine(current,time(12,0)))
                ax0.axvline(d_mid,color="tomato",lw=0.8,ls="--",alpha=0.5)
            current += timedelta(days=1)

    ax0.set_ylabel("BG (mmol/L)")
    ax0.set_title(profile_name+" | TDD="+str(tdd_info["tdd_avg"])+" U/day | "+
                  str(start_dt.date())+" to "+str(end_dt.date()),fontsize=10)
    ax0.legend(loc="upper right",fontsize=6,ncol=3)
    ax0.xaxis.set_major_locator(mdates.AutoDateLocator())
    ax0.xaxis.set_major_formatter(mdates.AutoDateFormatter(mdates.AutoDateLocator()))
    plt.setp(ax0.xaxis.get_majorticklabels(),rotation=20,fontsize=6)

    if oura_context:
        dates_list = sorted(oura_context.keys())
        dates_ts   = [pd.Timestamp(datetime.combine(d,time(12,0))) for d in dates_list]
        r_scores   = [float(oura_context[d].get("readiness_score") or 0) for d in dates_list]
        t_devs     = [float(oura_context[d].get("temperature_deviation") or 0) for d in dates_list]
        bar_colors = []
        for s in r_scores:
            bar_colors.append("#2ecc71" if s>=75 else "#a8d08d" if s>=65 else
                              "#f9e04b" if s>=55 else "#f0a500" if s>=45 else "#e74c3c")
        ax1.bar(dates_ts,r_scores,width=pd.Timedelta(hours=20),
                color=bar_colors,alpha=0.85,label="Readiness")
        ax1.axhline(70,color="gray",lw=0.6,ls="--",alpha=0.5)
        ax1.axhline(50,color="red", lw=0.6,ls="--",alpha=0.5)
        ax1.set_ylabel("Readiness",fontsize=7); ax1.set_ylim(0,100)
        ax1b = ax1.twinx()
        ax1b.plot(dates_ts,t_devs,color="tomato",lw=1.2,marker=".",markersize=4,
                  label="Temp dev (C)")
        ax1b.axhline(TEMP_DEV_FLAG, color="tomato",lw=0.5,ls=":")
        ax1b.axhline(-TEMP_DEV_FLAG,color="tomato",lw=0.5,ls=":")
        ax1b.set_ylabel("Temp dev (C)",fontsize=7,color="tomato")
        ax1b.tick_params(axis="y",labelcolor="tomato",labelsize=6)
        ax1.set_title("Oura: Readiness + Temperature",fontsize=8)
        ax1.xaxis.set_major_locator(mdates.AutoDateLocator())
        ax1.xaxis.set_major_formatter(mdates.AutoDateFormatter(mdates.AutoDateLocator()))
        plt.setp(ax1.xaxis.get_majorticklabels(),rotation=20,fontsize=6)
    else:
        ax1.text(0.5,0.5,"No Oura data",transform=ax1.transAxes,
                 ha="center",va="center",color="gray")

    valid = tech_df[tech_df["Basal Profile"].notna()&~tech_df["Basal Implausible"]].copy()
    if not valid.empty:
        seg_labels = valid["Interval"].tolist()
        x = np.arange(len(seg_labels)); w = 0.35
        now_v = valid["Basal Profile"].values.astype(float)
        rec_v = valid["Basal Rec Safe"].values
        rec_v = np.where(pd.isna(rec_v),now_v,rec_v.astype(float))
        ax2.bar(x-w/2,now_v,w,label="Current",    color="steelblue",alpha=0.8)
        ax2.bar(x+w/2,rec_v,w,label="Recommended",color="tomato",   alpha=0.8)
        ax2.set_xticks(x)
        ax2.set_xticklabels(seg_labels,rotation=45,ha="right",fontsize=6)
        ax2.set_ylabel("U/h",fontsize=7)
        ax2.set_title("Basal: Current vs Recommended (Safe)",fontsize=8)
        ax2.legend(fontsize=7)

    plt.tight_layout()
    return fig


def plot_attribution_diagnostic(cgm_aug, profile_name, start_dt, end_dt):
    colour_map = {"Basal":"steelblue","ISF":"darkorange","CSF":"forestgreen"}
    fig, axes  = plt.subplots(3,1,figsize=(16,10),sharex=True)
    ax0 = axes[0]
    for attr,col in colour_map.items():
        pts = cgm_aug[cgm_aug["attribution"]==attr]
        if not pts.empty:
            ax0.scatter(pts["Device Time"],pts["Value"],
                        c=col,s=4,label=attr,alpha=0.7,zorder=3)
    excl = cgm_aug[cgm_aug["meal_excluded"]==True]
    if not excl.empty:
        ax0.scatter(excl["Device Time"],excl["Value"],
                    c="gray",s=2,alpha=0.3,label="meal excluded",zorder=2)
    ax0.set_ylabel("BG (mmol/L)")
    ax0.set_title("Attribution | "+str(start_dt.date())+" to "+str(end_dt.date()))
    ax0.legend(markerscale=3,loc="upper right",fontsize=7)

    ax1 = axes[1]
    ax1.plot(cgm_aug["Device Time"],cgm_aug["bgi"],
             color="purple",lw=0.8,alpha=0.8,label="Insulin BGI")
    ax1.plot(cgm_aug["Device Time"],cgm_aug["abgi"],
             color="dodgerblue",lw=0.8,alpha=0.8,label="Activity BGI")
    ax1.plot(cgm_aug["Device Time"],cgm_aug["deviation"],
             color="crimson",lw=0.8,alpha=0.6,label="Deviation")
    ax1.axhline(0,color="black",lw=0.5,ls="--")
    ax1.axhline( DEV_NOISE_FLOOR,color="gray",lw=0.5,ls=":")
    ax1.axhline(-DEV_NOISE_FLOOR,color="gray",lw=0.5,ls=":")
    ax1.set_ylabel("mmol/L / 5 min"); ax1.legend(loc="upper right",fontsize=7)
    ax1.set_title("BGI components and Deviation")

    ax2 = axes[2]
    ax2.fill_between(cgm_aug["Device Time"],cgm_aug["cob"],
                     color="darkorange",alpha=0.45,label="COB (g)")
    if "in_workout" in cgm_aug.columns:
        wp = cgm_aug[cgm_aug["in_workout"]]
        if not wp.empty:
            ax2.scatter(wp["Device Time"],np.zeros(len(wp)),
                        c="dodgerblue",s=4,alpha=0.6,label="Workout")
    ax2.set_ylabel("COB (g)"); ax2.legend(loc="upper right",fontsize=7)
    ax2.set_title("COB and Workout periods")

    for ax in axes:
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        ax.xaxis.set_major_formatter(mdates.AutoDateFormatter(mdates.AutoDateLocator()))
    plt.setp(axes[-1].xaxis.get_majorticklabels(),rotation=25,fontsize=7)
    plt.tight_layout()
    plt.show()


def plot_meal_absorption(meals_enriched):
    if meals_enriched.empty or "absorption_score" not in meals_enriched.columns:
        print("No absorption data.")
        return
    df = meals_enriched.dropna(subset=["absorption_score"]).copy()
    if df.empty:
        print("No meals with absorption scores.")
        return
    color_map = {"fast":"tomato","normal":"steelblue","slow":"forestgreen","unknown":"gray"}
    fig, axes = plt.subplots(2,1,figsize=(13,6),sharex=True)
    ax0 = axes[0]
    for atype, grp in df.groupby("absorption_type"):
        ax0.scatter(grp["Timestamp"],grp["absorption_score"],
                    c=color_map.get(atype,"gray"),s=40,alpha=0.85,
                    label=atype.upper(),zorder=3)
    ax0.axhline(65,color="tomato",     lw=0.8,ls="--",label="Fast threshold (65)")
    ax0.axhline(40,color="forestgreen",lw=0.8,ls="--",label="Slow threshold (40)")
    ax0.set_ylabel("Absorption score (GI corrected)")
    ax0.set_title("Meal Absorption Speed  [fast=red  normal=blue  slow=green]",fontsize=9)
    ax0.legend(fontsize=7,ncol=5)
    ax1 = axes[1]
    watch = df[df.get("watch_matched",False)==True].dropna(subset=["bolus_delta_min"]) \
            if "watch_matched" in df.columns else pd.DataFrame()
    if not watch.empty:
        colors_w = [color_map.get(t,"gray") for t in watch["absorption_type"]]
        ax1.bar(watch["Timestamp"],watch["bolus_delta_min"],
                width=pd.Timedelta(minutes=20),color=colors_w,alpha=0.75)
        ax1.axhline(0, color="black", lw=0.8,ls="--")
        ax1.axhline(10,color="tomato",lw=0.6,ls=":",label=">10 min late")
        ax1.axhline(-5,color="green", lw=0.6,ls=":",label="5+ min pre-bolus")
        ax1.set_ylabel("Bolus delta (min)"); ax1.legend(fontsize=7)
        ax1.set_title("Pre-bolus timing (Watch data only)",fontsize=9)
    else:
        ax1.text(0.5,0.5,"No Watch timing data",transform=ax1.transAxes,
                 ha="center",color="gray")
    for ax in axes:
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        ax.xaxis.set_major_formatter(mdates.AutoDateFormatter(mdates.AutoDateLocator()))
    plt.setp(axes[-1].xaxis.get_majorticklabels(),rotation=25,fontsize=7)
    plt.tight_layout()
    plt.show()


def plot_basal_trend(basal_trend):
    df = basal_trend["daily_df"]
    if df.empty or len(df)<2:
        print("Not enough daily data.")
        return
    fig, ax = plt.subplots(figsize=(14,4))
    x = np.arange(len(df))
    ax.bar(x,df["basal_U"].values,color="steelblue",alpha=0.8,label="Daily basal (U)")
    if basal_trend["n_days"] >= MIN_DAYS_FOR_TREND:
        y = df["basal_U"].values.astype(float)
        slope,intercept,*_ = scipy_stats.linregress(x,y)
        ax.plot(x,intercept+slope*x,color="crimson",lw=2,
                label="Trend ("+_fmt(slope,2)+" U/day)")
    ax.set_xticks(x)
    ax.set_xticklabels([str(d) for d in df["date"].tolist()],
                       rotation=40,ha="right",fontsize=7)
    ax.set_ylabel("Delivered basal (U/day)")
    ax.set_title("Daily Basal Trend  |  CV="+_fmt(basal_trend["cv_pct"],1)+
                 "%  |  "+basal_trend["trend_label"].upper(),fontsize=10)
    ax.legend(fontsize=8)
    plt.tight_layout()
    plt.show()


# ================================================================
# TABLE STYLING
# ================================================================

def style_action_table(df):
    col_widths = {
        "Segment":"78px","Status":"68px","Mean BG":"52px","TIR %":"44px",
        "CV %":"44px","Cap %":"38px","Absorption":"58px","Act pts":"44px",
        "Basal now":"55px","Basal ->":"55px","CR now":"48px","CR ->":"48px",
        "ISF now":"48px","ISF ->":"48px","K1 now":"44px","K1 ->":"44px",
        "Confidence":"68px","Action":"260px",
    }
    col_styles = []
    for col, w in col_widths.items():
        if col not in df.columns: continue
        idx = df.columns.get_loc(col)
        nth = str(idx+1)
        col_styles.append({"selector":"th:nth-child("+nth+")",
            "props":"width:"+w+";min-width:"+w+";max-width:"+w+";"
                    "text-align:center;font-weight:bold;font-size:9px;"
                    "white-space:nowrap;overflow:hidden;padding:3px;"})
        if col == "Action":
            col_styles.append({"selector":"td:nth-child("+nth+")",
                "props":"width:"+w+";min-width:160px;max-width:"+w+";"
                        "text-align:left;font-size:8px;padding:3px;"
                        "white-space:normal;word-wrap:break-word;"
                        "overflow-wrap:break-word;vertical-align:top;"})
        else:
            col_styles.append({"selector":"td:nth-child("+nth+")",
                "props":"width:"+w+";max-width:"+w+";"
                        "text-align:center;font-size:9px;padding:3px;"
                        "white-space:nowrap;overflow:hidden;vertical-align:middle;"})
    base_styles = [{"selector":"table",
                    "props":"table-layout:fixed;width:100%;border-collapse:collapse;"}]

    def highlight(row):
        out = []
        status = str(row.get("Status",""))
        for col in row.index:
            if col=="Status":
                if status=="HYPER":   out.append("background-color:#ffcccc;font-weight:bold")
                elif status=="HYPO":  out.append("background-color:#ffd580;font-weight:bold")
                elif status=="OPTIMAL": out.append("background-color:#c8f5c8;font-weight:bold")
                else:                 out.append("background-color:#fffacd")
            elif col=="Confidence":
                v = str(row[col])
                if v=="High":  out.append("background-color:#c8f5c8")
                elif v=="Low": out.append("background-color:#ffe4b5")
                else:          out.append("")
            elif col in ("Basal ->","CR ->","ISF ->","K1 ->"): out.append("background-color:#fff2cc")
            elif col in ("Basal now","CR now","ISF now","K1 now"): out.append("background-color:#e6f9e6")
            elif col in ("Cap %","Act pts"): out.append("background-color:#f0f0f0")
            elif col=="Absorption":
                v = str(row[col])
                if v=="fast":   out.append("background-color:#fde8e8")
                elif v=="slow": out.append("background-color:#e8f4e8")
                else:           out.append("")
            elif col=="Action":
                v = str(row[col])
                if v.startswith("WARNING"):
                    out.append("background-color:#fff0f0;color:#c00000;font-weight:bold;")
                else: out.append("")
            else: out.append("")
        return out

    return (df.style.hide(axis="index")
              .set_table_styles(base_styles+col_styles,overwrite=True)
              .apply(highlight,axis=1))


# ================================================================
# PDF EXPORT
# ================================================================

def export_pdf(summary_text, action_df, chart_fig, tdd_info,
               profile_name, start_dt, end_dt, pdf_path):
    buf = io.BytesIO()
    chart_fig.savefig(buf,format="png",dpi=100,bbox_inches="tight")
    buf.seek(0)
    c  = canvas.Canvas(pdf_path, pagesize=letter)
    W, H = letter
    margin = 28
    c.setFont("Helvetica-Bold",11)
    c.drawString(margin,H-33,"AUTOTUNE v5 -- "+profile_name+
                 "  |  "+str(start_dt.date())+" to "+str(end_dt.date()))
    chart_h = 230; chart_y = H-48-chart_h
    c.drawImage(ImageReader(PILImage.open(buf)),margin,chart_y,
                width=W-2*margin,height=chart_h)
    text_y = chart_y-10
    for line in summary_text.split("\n"):
        if text_y < 188: break
        c.setFont("Helvetica-Bold",7) if (line.isupper() and len(line)>3) else c.setFont("Helvetica",6.5)
        c.drawString(margin,text_y,line)
        text_y -= 7.5
    if not action_df.empty:
        pdf_cols = ["Segment","Status","Mean BG","TIR %","Cap %","Absorption",
                    "Basal now","Basal ->","CR now","CR ->",
                    "ISF now","ISF ->","K1 now","K1 ->","Confidence","Action"]
        ecols   = [col for col in pdf_cols if col in action_df.columns]
        pdf_act = action_df[ecols].fillna("-")
        tdata   = [list(pdf_act.columns)]
        for _, row in pdf_act.iterrows():
            tdata.append([str(v) for v in row.values])
        tbl = Table(tdata,repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.Color(0.2,0.4,0.6)),
            ("TEXTCOLOR",(0,0),(-1,0),colors.whitesmoke),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),4.8),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.Color(0.95,0.95,0.95)]),
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ]))
        tbl.wrapOn(c,W-2*margin,160)
        tbl.drawOn(c,margin,margin)
    c.save()
    print("PDF saved: "+pdf_path)


# ================================================================
# USER INTERFACE
# ================================================================

basal_df_ui,_,_,_,_,_ = load_tidepool_data(FILE_PATH)
profile_names_ui = basal_df_ui["Schedule Name"].unique()
_cfg = load_config()

profile_dd    = widgets.Dropdown(options=profile_names_ui,description="Profile:")
start_dp      = widgets.DatePicker(description="Start date")
end_dp        = widgets.DatePicker(description="End date")
sleep_start_w = widgets.Text(value=_cfg.get("sleep_start","23:00"),description="Sleep start:")
sleep_end_w   = widgets.Text(value=_cfg.get("sleep_end","07:00"),  description="Sleep end:")
extended_w    = widgets.Checkbox(value=False,description="Extended mode")
btn = widgets.Button(description="Run Autotune v5",button_style="primary",
                     layout=widgets.Layout(width="230px"))
out = widgets.Output()


def on_click(b):
    with out:
        clear_output()
        if not profile_dd.value or not start_dp.value or not end_dp.value:
            print("Please select a profile and both dates.")
            return

        cfg = load_config()
        cfg["sleep_start"] = sleep_start_w.value
        cfg["sleep_end"]   = sleep_end_w.value
        save_config(cfg)

        s_dt = datetime.combine(start_dp.value,datetime.min.time())
        e_dt = datetime.combine(end_dp.value,  datetime.max.time())

        if (e_dt-s_dt).days < 7:
            print("WARNING: less than 7 days. Recommend >= 14.")

        print("Running Autotune v5 Final ...")

        try:
            # Oura
            print("  [1/5] Loading Oura data ...")
            try:
                access_token = get_oura_access_token()
                oura_data    = load_oura_data(access_token,s_dt,e_dt)
                oura_context = preprocess_oura(oura_data,s_dt,e_dt)
                print("  Oura: OK")
            except Exception as e:
                print("  Oura: FAILED ("+str(e)+") -- continuing without")
                oura_context = {}

            # Apple Watch workouts
            print("  [2/5] Loading Apple Watch workouts ...")
            workouts_df      = load_watch_workouts(WATCH_CSV_PATH,s_dt,e_dt)
            activity_windows = build_activity_windows(workouts_df)
            print("  Watch workouts: "+str(len(activity_windows))+" loaded")

            # Meal data
            print("  [3/5] Loading meal data ...")
            dnevnik_df  = parse_dnevnik_xml(DNEVNIK_XML_PATH,s_dt,e_dt)
            watch_pairs = parse_meal_log(MEAL_LOG_PATH,s_dt,e_dt)
            meals_raw   = load_bolus_calculator(FILE_PATH)
            meals_raw   = meals_raw[
                (meals_raw["Timestamp"]>=s_dt)&
                (meals_raw["Timestamp"]<=e_dt)
            ].reset_index(drop=True)
            meals_enriched = enrich_meals(meals_raw,dnevnik_df,watch_pairs)
            timing_summary = analyze_bolus_timing(meals_enriched)

            # Main analysis
            print("  [4/5] Running main analysis ...")
            tech_df,cgm_aug,tdd_info,overall,basal_trend,cal_factor = run_autotune(
                profile_dd.value,s_dt,e_dt,
                sleep_start_w.value,sleep_end_w.value,
                oura_context,activity_windows,meals_enriched)

            # Output
            print("  [5/5] Building output ...")
            action_df = build_action_table(tech_df)
            summary   = build_summary(tech_df,tdd_info,overall,basal_trend,
                                      oura_context,meals_enriched,timing_summary,
                                      profile_dd.value,s_dt,e_dt)
            chart_fig = build_summary_chart(cgm_aug,tech_df,tdd_info,oura_context,
                                            activity_windows,profile_dd.value,s_dt,e_dt)

            print("="*68)
            print(summary)
            print("="*68)
            print("")
            print("ACTION TABLE")
            if action_df.empty:
                print("No segments with sufficient data.")
            else:
                display(style_action_table(action_df))
            print("")
            plt.show()

            if extended_w.value:
                print("\nEXTENDED MODE")
                print("-"*60)
                plot_basal_trend(basal_trend)
                plot_meal_absorption(meals_enriched)
                plot_attribution_diagnostic(cgm_aug,profile_dd.value,s_dt,e_dt)

            pdf_path = os.path.join(BASE_PATH,"autotune_v5_final.pdf")
            export_pdf(summary,action_df,chart_fig,tdd_info,
                       profile_dd.value,s_dt,e_dt,pdf_path)

        except Exception as exc:
            import traceback
            print("ERROR: "+str(exc))
            traceback.print_exc()


btn.on_click(on_click)
display(profile_dd,start_dp,end_dp,sleep_start_w,sleep_end_w,extended_w,btn,out)
