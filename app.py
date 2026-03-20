import streamlit as st
import pandas as pd
import re
import io
from urllib.parse import unquote, quote
from difflib import SequenceMatcher
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Lead Classifier",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# MARKET CONFIG
# ─────────────────────────────────────────────
MARKETS = {
    "TR": {
        "name": "Turkey",
        "flag": "🇹🇷",
        "char_map": {
            "Ç": "C", "Ğ": "G", "İ": "I", "Ö": "O", "Ş": "S", "Ü": "U",
            "ç": "c", "ğ": "g", "ı": "i", "ö": "o", "ş": "s", "ü": "u",
        },
        "country_suffix": "Turkey",
        "phone_prefix": "90",
    },
    "NO": {
        "name": "Norway",
        "flag": "🇳🇴",
        "char_map": {
            "Æ": "AE", "Ø": "O", "Å": "A",
            "æ": "ae", "ø": "o", "å": "a",
        },
        "country_suffix": "Norway",
        "phone_prefix": "47",
    },
    "SE": {
        "name": "Sweden",
        "flag": "🇸🇪",
        "char_map": {
            "Å": "A", "Ä": "A", "Ö": "O",
            "å": "a", "ä": "a", "ö": "o",
        },
        "country_suffix": "Sweden",
        "phone_prefix": "46",
    },
    "CZ": {
        "name": "Czech Republic",
        "flag": "🇨🇿",
        "char_map": {
            "Á": "A", "Č": "C", "Ď": "D", "É": "E", "Ě": "E", "Í": "I",
            "Ň": "N", "Ó": "O", "Ř": "R", "Š": "S", "Ť": "T", "Ú": "U",
            "Ů": "U", "Ý": "Y", "Ž": "Z",
            "á": "a", "č": "c", "ď": "d", "é": "e", "ě": "e", "í": "i",
            "ň": "n", "ó": "o", "ř": "r", "š": "s", "ť": "t", "ú": "u",
            "ů": "u", "ý": "y", "ž": "z",
        },
        "country_suffix": "Czech Republic",
        "phone_prefix": "420",
    },
    "HU": {
        "name": "Hungary",
        "flag": "🇭🇺",
        "char_map": {
            "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ö": "O", "Ő": "O",
            "Ú": "U", "Ü": "U", "Ű": "U",
            "á": "a", "é": "e", "í": "i", "ó": "o", "ö": "o", "ő": "o",
            "ú": "u", "ü": "u", "ű": "u",
        },
        "country_suffix": "Hungary",
        "phone_prefix": "36",
    },
    "AT": {
        "name": "Austria",
        "flag": "🇦🇹",
        "char_map": {
            "Ä": "A", "Ö": "O", "Ü": "U",
            "ä": "a", "ö": "o", "ü": "u",
        },
        "country_suffix": "Austria",
        "phone_prefix": "43",
    },
}

NON_FOOD_CATEGORIES = {
    "Haute couture fashion house", "Tattoo shop", "Hotel", "Psychologist",
    "Corporate office", "Boutique", "Shopping mall", "Digital printer",
    "Industrial equipment supplier", "Bakery equipment", "Bus stop",
    "Military residence", "Chess club", "Event planner", "Betting agency",
    "Home goods store", "Natural goods store", "General store",
    "Convenience store", "Food manufacturer", "Food manufacturing supply",
    "Beach club", "Wedding venue", "Rest stop", "Hair salon", "Beauty salon",
    "Gym", "Fitness center", "Pharmacy", "Dentist", "Doctor", "Hospital",
    "Bank", "Insurance agency", "Car dealership", "Auto repair shop",
    "Clothing store", "Electronics store", "Furniture store",
}

# ─────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────
def norm_name(s, char_map):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    for k, v in char_map.items():
        s = s.replace(k, v)
    return s.lower().strip()


def norm_phone(p, prefix):
    if pd.isna(p):
        return ""
    s = str(p).replace("+", "").replace(" ", "").replace("-", "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.lstrip("0")
    if not s.startswith(prefix):
        s = prefix + s
    return s


def to_e164(p, prefix):
    n = norm_phone(p, prefix)
    return "+" + n if n else ""


def norm_url(u):
    if pd.isna(u):
        return ""
    return unquote(str(u).strip()).split("?hl=")[0].lower()


def name_confidence(a, b, char_map):
    a = norm_name(a, char_map)
    b = norm_name(b, char_map)
    if not a or not b:
        return 0.0
    tokens_a = set(re.split(r"\W+", a)) - {""}
    tokens_b = set(re.split(r"\W+", b)) - {""}
    if not tokens_a or not tokens_b:
        return 0.0
    shorter = tokens_a if len(tokens_a) <= len(tokens_b) else tokens_b
    token_score = len(shorter & tokens_b) / len(shorter)
    seq_score = SequenceMatcher(None, a, b).ratio()
    return round(max(token_score, seq_score), 3)


def address_match(lead_street, apify_address, char_map):
    if pd.isna(lead_street) or pd.isna(apify_address):
        return False
    ls = norm_name(str(lead_street), char_map)
    aa = norm_name(str(apify_address), char_map)
    postal_l = set(re.findall(r"\b\d{4,5}\b", ls))
    postal_a = set(re.findall(r"\b\d{4,5}\b", aa))
    if postal_l and postal_a and postal_l & postal_a:
        return True
    ignore = {"", "no", "sk", "cd", "mh", "tr", "ve", "de", "da", "og", "och"}
    tokens_l = set(re.split(r"\W+", ls)) - ignore
    tokens_a = set(re.split(r"\W+", aa)) - ignore
    return len(tokens_l & tokens_a) >= 2


def detect_column(df, candidates):
    """Find the first column name from candidates that exists in df."""
    for c in candidates:
        if c in df.columns:
            return c
    # case-insensitive fallback
    lower_map = {col.lower(): col for col in df.columns}
    for c in candidates:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    return None


def load_leads(file, market_cfg):
    """Load leads file and auto-detect columns."""
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        # Try different header rows for xlsx
        for header_row in [0, 5, 10, 14]:
            try:
                df = pd.read_excel(file, header=header_row)
                if len(df.columns) > 3 and df.shape[0] > 5:
                    break
            except Exception:
                continue

    col_map = {}
    col_map["name"]    = detect_column(df, ["Company / Account", "Account Name", "Name", "company_name", "Företag", "Virksomhed", "Vállalkozás", "Unternehmen"])
    col_map["phone"]   = detect_column(df, ["Phone", "phone_number", "Telefon", "Telefonnummer"])
    col_map["street"]  = detect_column(df, ["Street", "Address", "address", "Adresse", "Cím"])
    col_map["city"]    = detect_column(df, ["City", "city", "By", "Stad", "Město", "Város", "Stadt"])
    col_map["grid"]    = detect_column(df, ["GRID", "grid", "Grid"])
    col_map["lead_id"] = detect_column(df, ["Lead ID", "lead_id", "LeadID", "ID"])
    col_map["url"]     = detect_column(df, ["GOOGLE URL", "Google URL", "google_url", "URL"])
    col_map["lat"]     = detect_column(df, ["Coordinates (Latitude)", "Latitude", "lat"])
    col_map["lng"]     = detect_column(df, ["Coordinates (Longitude)", "Longitude", "lng"])

    return df, col_map


def load_crm(file, market_cfg):
    """Load CRM file, handling Salesforce report headers."""
    prefix = market_cfg["phone_prefix"]
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        for header_row in [0, 5, 10, 14]:
            try:
                df = pd.read_excel(file, header=header_row, skipfooter=2)
                if len(df.columns) > 3 and df.shape[0] > 5:
                    break
            except Exception:
                continue

    col_map = {}
    col_map["grid"]   = detect_column(df, ["GRID__c", "GRID", "Grid"])
    col_map["name"]   = detect_column(df, ["Account Name", "Name", "name"])
    col_map["phone"]  = detect_column(df, ["Phone", "phone"])
    col_map["status"] = detect_column(df, ["Account_Status__c", "Account Status", "AccountStatus"])
    col_map["reason"] = detect_column(df, ["Status_Reason__c", "Status Reason", "StatusReason"])

    phone_col = col_map["phone"]
    if phone_col:
        df["_phone_norm"] = df[phone_col].apply(lambda p: norm_phone(p, prefix))

    name_col = col_map["name"]
    char_map = market_cfg["char_map"]
    if name_col:
        df["_name_norm"] = df[name_col].apply(lambda n: norm_name(n, char_map))

    return df, col_map


def load_apify(file):
    """Load Apify results file."""
    if file.name.endswith(".csv"):
        df = pd.read_csv(file, low_memory=False)
    else:
        df = pd.read_excel(file)

    col_map = {}
    col_map["title"]    = detect_column(df, ["title"])
    col_map["url"]      = detect_column(df, ["searchPageUrl"])
    col_map["gm_url"]   = detect_column(df, ["url"])
    col_map["phone"]    = detect_column(df, ["phone"])
    col_map["website"]  = detect_column(df, ["website"])
    col_map["category"] = detect_column(df, ["categoryName"])
    col_map["perm"]     = detect_column(df, ["permanentlyClosed"])
    col_map["temp"]     = detect_column(df, ["temporarilyClosed"])
    col_map["address"]  = detect_column(df, ["address"])

    if col_map["url"]:
        df["_url_norm"] = df[col_map["url"]].apply(norm_url)

    return df, col_map


def generate_google_urls(leads_df, col_map, market_cfg):
    """Generate Google Maps search URLs for each lead."""
    name_col   = col_map.get("name")
    street_col = col_map.get("street")
    lat_col    = col_map.get("lat")
    lng_col    = col_map.get("lng")
    suffix     = market_cfg["country_suffix"]

    urls = []
    for _, row in leads_df.iterrows():
        name   = str(row[name_col]).strip()   if name_col   and pd.notna(row.get(name_col))   else ""
        street = str(row[street_col]).strip() if street_col and pd.notna(row.get(street_col)) else ""
        lat    = row.get(lat_col)             if lat_col    else None
        lng    = row.get(lng_col)             if lng_col    else None

        if lat and lng and pd.notna(lat) and pd.notna(lng):
            # Coordinate URL (most accurate)
            url = f"https://www.google.com/maps/search/?api=1&query={lat},{lng}"
        elif name and street:
            # Name + address URL
            query = quote(f"{name} {street}")
            url = f"https://www.google.com/maps/search/{query}"
        elif name:
            query = quote(f"{name} {suffix}")
            url = f"https://www.google.com/maps/search/{query}"
        else:
            url = ""
        urls.append(url)

    return urls


def classify_leads(leads_df, col_map_leads, crm_df, col_map_crm,
                   apify_df, col_map_apify, market_cfg,
                   confidence_threshold=0.5):
    """Main classification pipeline."""
    char_map = market_cfg["char_map"]
    prefix   = market_cfg["phone_prefix"]

    # ── Build CRM lookups ──────────────────────────────────────────
    crm_phone_dict = {}
    crm_name_dict  = {}

    if crm_df is not None:
        phone_col  = col_map_crm.get("phone")
        name_col   = col_map_crm.get("name")
        grid_col   = col_map_crm.get("grid")
        status_col = col_map_crm.get("status")
        reason_col = col_map_crm.get("reason")

        for _, r in crm_df.iterrows():
            pn = r.get("_phone_norm", "")
            if pn and pn not in crm_phone_dict:
                crm_phone_dict[pn] = r
            nn = r.get("_name_norm", "")
            if nn and nn not in crm_name_dict:
                crm_name_dict[nn] = r

    # ── Build Apify lookup ─────────────────────────────────────────
    apify_dict = {}
    if apify_df is not None and "_url_norm" in apify_df.columns:
        for _, r in apify_df.drop_duplicates("_url_norm").iterrows():
            apify_dict[r["_url_norm"]] = r

    # ── Classify each lead ─────────────────────────────────────────
    name_col_l   = col_map_leads.get("name")
    phone_col_l  = col_map_leads.get("phone")
    street_col_l = col_map_leads.get("street")
    city_col_l   = col_map_leads.get("city")
    grid_col_l   = col_map_leads.get("grid")
    lead_id_col  = col_map_leads.get("lead_id")
    url_col_l    = col_map_leads.get("url")

    results = []
    for _, row in leads_df.iterrows():
        phone = norm_phone(row.get(phone_col_l, ""), prefix) if phone_col_l else ""
        lead_name = row.get(name_col_l, "") if name_col_l else ""
        lead_street = row.get(street_col_l, "") if street_col_l else ""

        # URL key: use explicit URL column if present, otherwise reconstruct from name+street
        # (matches what the URL generator tab produces and what Apify searchPageUrl contains)
        if url_col_l and pd.notna(row.get(url_col_l, "")):
            url = norm_url(row.get(url_col_l, ""))
        else:
            parts = str(lead_name).strip()
            if lead_street and not pd.isna(lead_street):
                parts += " " + str(lead_street).strip()
            from urllib.parse import quote as _quote
            url = norm_url("https://www.google.com/maps/search/" + _quote(parts)) if parts else ""

        # ── CRM duplicate check ────────────────────────────────────
        crm_match    = crm_phone_dict.get(phone) if phone else None
        match_method = "Phone" if crm_match is not None else ""
        if crm_match is None:
            nn = norm_name(lead_name, char_map)
            crm_match = crm_name_dict.get(nn) if nn else None
            if crm_match is not None:
                match_method = "Name"

        label = dup_grid = dup_name = dup_status = dup_reason = dup_method = ""
        if crm_match is not None:
            label      = "Duplicate"
            dup_grid   = str(crm_match.get(col_map_crm.get("grid", ""), ""))
            dup_name   = str(crm_match.get(col_map_crm.get("name", ""), ""))
            dup_status = str(crm_match.get(col_map_crm.get("status", ""), ""))
            dup_reason = str(crm_match.get(col_map_crm.get("reason", ""), "") or "")
            dup_method = match_method

        # ── Apify enrichment + validation ─────────────────────────
        apy = apify_dict.get(url) if url else None
        gm_title = gm_cat = gm_biz_status = gm_phone = gm_website = gm_url = ""
        match_conf = 0.0
        match_reason = ""

        if apy is not None:
            tc  = col_map_apify.get("title");    gm_title   = str(apy[tc]  or "") if tc  and pd.notna(apy.get(tc))  else ""
            cc  = col_map_apify.get("category"); gm_cat     = str(apy[cc]  or "") if cc  and pd.notna(apy.get(cc))  else ""
            pc  = col_map_apify.get("phone");    gm_phone   = to_e164(apy.get(pc, ""), prefix) if pc else ""
            wc  = col_map_apify.get("website");  gm_website = str(apy[wc]  or "") if wc  and pd.notna(apy.get(wc))  else ""
            uc  = col_map_apify.get("gm_url");   gm_url     = str(apy[uc]  or "") if uc  and pd.notna(apy.get(uc))  else ""
            permc = col_map_apify.get("perm");   perm = str(apy.get(permc, "")).lower() == "true" if permc else False
            tempc = col_map_apify.get("temp");   temp = str(apy.get(tempc, "")).lower() == "true" if tempc else False

            if perm:   gm_biz_status = "Permanently Closed"
            elif temp: gm_biz_status = "Temporarily Closed"
            else:      gm_biz_status = "Open"

            gm_phone_norm = norm_phone(apy.get(pc, ""), prefix) if pc else ""
            phone_hit  = bool(phone and phone == gm_phone_norm)
            name_score = name_confidence(lead_name, gm_title, char_map)
            addr_hit   = address_match(
                row.get(street_col_l, ""), apy.get(col_map_apify.get("address", ""), ""), char_map
            )
            confirmed = phone_hit or (name_score >= confidence_threshold)

            reasons = []
            if phone_hit: reasons.append("Phone ✓")
            reasons.append(f"Name {name_score:.2f}" + (" ✓" if name_score >= confidence_threshold else " ✗"))
            if addr_hit:  reasons.append("Address ✓")
            match_reason = " | ".join(reasons)
            match_conf   = round(
                (0.4 if phone_hit else 0.0) +
                (min(name_score, 1.0) * 0.4) +
                (0.2 if addr_hit else 0.0), 3
            )

            if not label:
                if not confirmed:                   label = "Invalid Data"
                elif not gm_cat:                    label = "Invalid Data"
                elif perm or temp:                  label = "Business Closed"
                elif gm_cat in NON_FOOD_CATEGORIES: label = "Wrong Target Group"
                else:                               label = "Qualified / Convert"
        else:
            gm_biz_status = "Not Found on Google"
            match_reason  = "No Apify result"
            if not label:
                label = "Invalid Data"

        results.append({
            "GRID":                  row.get(grid_col_l, "")    if grid_col_l    else "",
            "Lead ID":               row.get(lead_id_col, "")   if lead_id_col   else "",
            "Company / Account":     lead_name,
            "City":                  row.get(city_col_l, "")    if city_col_l    else "",
            "Street":                row.get(street_col_l, "")  if street_col_l  else "",
            "Phone":                 to_e164(row.get(phone_col_l, ""), prefix) if phone_col_l else "",
            "GM Title":              gm_title,
            "GM Category":           gm_cat,
            "GM Business Status":    gm_biz_status,
            "GM Phone":              gm_phone,
            "GM Website":            gm_website,
            "GM URL":                gm_url,
            "Match Confidence":      match_conf,
            "Match Reason":          match_reason,
            "Label":                 label,
            "Duplicate GRID":        dup_grid,
            "Duplicate CRM Name":    dup_name,
            "CRM Account Status":    dup_status,
            "CRM Status Reason":     dup_reason,
            "Duplicate Match Method":dup_method,
        })

    return pd.DataFrame(results)


# ─────────────────────────────────────────────
# EXCEL OUTPUT
# ─────────────────────────────────────────────
def build_excel(df, market_name):
    FILLS = {
        "Qualified / Convert": PatternFill("solid", start_color="C6EFCE"),
        "Duplicate":           PatternFill("solid", start_color="FFC7CE"),
        "Business Closed":     PatternFill("solid", start_color="FFEB9C"),
        "Wrong Target Group":  PatternFill("solid", start_color="FFDCA8"),
        "Invalid Data":        PatternFill("solid", start_color="D9D9D9"),
    }
    ALT_FILLS = {
        "Qualified / Convert": PatternFill("solid", start_color="EAF5EA"),
        "Duplicate":           PatternFill("solid", start_color="FFE0E0"),
        "Business Closed":     PatternFill("solid", start_color="FFF7D1"),
        "Wrong Target Group":  PatternFill("solid", start_color="FFF0DC"),
        "Invalid Data":        PatternFill("solid", start_color="EFEFEF"),
    }
    LABEL_COLORS = {
        "Qualified / Convert": "276221",
        "Duplicate":           "9C0006",
        "Business Closed":     "7D4E00",
        "Wrong Target Group":  "833C00",
        "Invalid Data":        "595959",
    }
    HEADER_FILL  = PatternFill("solid", start_color="1F4E79")
    SECTION_FILL = PatternFill("solid", start_color="2E75B6")

    def thin():
        s = Side(style="thin", color="D0D0D0")
        return Border(left=s, right=s, top=s, bottom=s)

    def hdr(ws, row, col, val):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin()
        return c

    def sec(ws, row, col, text, span=3):
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
        c = ws.cell(row=row, column=col, value=f"  {text}")
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.fill = SECTION_FILL
        c.alignment = Alignment(vertical="center")
        ws.row_dimensions[row].height = 22

    def dc(ws, row, col, val, fill=None, bold=False, fmt=None, align="center", color="000000"):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name="Arial", size=10, bold=bold, color=color)
        if fill:
            c.fill = fill
        c.alignment = Alignment(horizontal=align, vertical="center")
        c.border = thin()
        if fmt:
            c.number_format = fmt
        return c

    def conf_fill(score):
        try:
            score = float(score)
        except Exception:
            return PatternFill()
        if score >= 0.8: return PatternFill("solid", start_color="C6EFCE")
        if score >= 0.6: return PatternFill("solid", start_color="FFEB9C")
        return PatternFill("solid", start_color="FFC7CE")

    wb = Workbook()
    counts = df["Label"].value_counts()
    labels_order = ["Qualified / Convert", "Duplicate", "Business Closed", "Wrong Target Group", "Invalid Data"]
    DATA_S = 5
    DATA_E = DATA_S + len(df) - 1

    LBL_R  = f"'Classified Leads'!O{DATA_S}:O{DATA_E}"
    CAT_R  = f"'Classified Leads'!H{DATA_S}:H{DATA_E}"
    CITY_R = f"'Classified Leads'!D{DATA_S}:D{DATA_E}"
    CRMS_R = f"'Classified Leads'!Q{DATA_S}:Q{DATA_E}"
    METH_R = f"'Classified Leads'!T{DATA_S}:T{DATA_E}"

    col_headers = [
        "GRID", "Lead ID", "Company / Account", "City", "Street", "Phone",
        "GM Title", "GM Category", "GM Business Status",
        "GM Phone", "GM Website", "GM URL",
        "Match Confidence", "Match Reason",
        "Label",
        "Duplicate GRID", "Duplicate CRM Name", "CRM Account Status", "CRM Status Reason",
        "Duplicate Match Method",
    ]

    # ── Sheet 1: All leads ───────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Classified Leads"
    ws1["A1"] = f"Lead Classification Report  |  {market_name}"
    ws1["A1"].font = Font(name="Arial", bold=True, size=14, color="1F4E79")
    ws1.merge_cells("A1:T1")
    ws1["A2"] = "Total: {:,}   |   {}".format(
        len(df),
        "   ".join(f"{l}: {counts.get(l, 0):,}" for l in labels_order),
    )
    ws1["A2"].font = Font(name="Arial", italic=True, size=9, color="595959")
    ws1.merge_cells("A2:T2")

    for ci, h in enumerate(col_headers, 1):
        hdr(ws1, 4, ci, h)
    ws1.row_dimensions[4].height = 30

    for ri, (_, row) in enumerate(df.iterrows(), DATA_S):
        lbl  = row["Label"]
        fill = FILLS.get(lbl, PatternFill()) if ri % 2 == 0 else ALT_FILLS.get(lbl, PatternFill())
        lc   = LABEL_COLORS.get(lbl, "000000")
        for ci, key in enumerate(col_headers, 1):
            val = row.get(key, "")
            if pd.isna(val):
                val = ""
            c = ws1.cell(row=ri, column=ci, value=val if val != "" else "")
            c.border = thin()
            if key == "Label":
                c.font = Font(name="Arial", size=10, bold=True, color=lc)
                c.fill = fill
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif key == "Match Confidence":
                try:
                    c.value = float(val)
                    c.number_format = "0.00"
                except Exception:
                    pass
                c.font = Font(name="Arial", size=10, bold=True)
                c.fill = conf_fill(val)
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif key == "Match Reason":
                c.font = Font(name="Arial", size=9, color="595959")
                c.fill = fill
                c.alignment = Alignment(vertical="center")
            else:
                c.font = Font(name="Arial", size=10)
                c.fill = fill
                c.alignment = Alignment(vertical="center")

    col_widths = [10, 18, 32, 14, 36, 16, 30, 24, 18, 16, 30, 50, 14, 38, 22, 14, 30, 16, 18, 16]
    for i, w in enumerate(col_widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = "A5"
    ws1.auto_filter.ref = f"A4:T{DATA_E}"

    # ── Sheet 2: Summary ─────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Classification Summary"
    ws2["A1"].font = Font(name="Arial", bold=True, size=16, color="1F4E79")
    ws2["A2"] = 'All counts are formula-driven and update automatically when "Classified Leads" is edited.'
    ws2["A2"].font = Font(name="Arial", italic=True, size=9, color="595959")
    ws2.merge_cells("A2:J2")

    sec(ws2, 4, 1, "CLASSIFICATION BREAKDOWN", span=3)
    for ci, h in enumerate(["Label", "Count", "% of Total"], 1):
        hdr(ws2, 5, ci, h)
    for i, lbl in enumerate(labels_order):
        r = 6 + i
        dc(ws2, r, 1, lbl, fill=FILLS.get(lbl), bold=True, align="left", color=LABEL_COLORS.get(lbl, "000000"))
        dc(ws2, r, 2, f'=COUNTIF({LBL_R},"{lbl}")', fill=FILLS.get(lbl))
        dc(ws2, r, 3, f"=IF(B11=0,0,B{r}/B11)", fill=FILLS.get(lbl), fmt="0.0%")
    dc(ws2, 11, 1, "TOTAL", bold=True, align="left")
    dc(ws2, 11, 2, "=SUM(B6:B10)", bold=True)
    dc(ws2, 11, 3, "100.0%", bold=True, fmt="0.0%")

    sec(ws2, 13, 1, "DUPLICATE MATCH METHOD", span=3)
    for ci, h in enumerate(["Method", "Count", "% of Dupes"], 1):
        hdr(ws2, 14, ci, h)
    for i, (method, fill) in enumerate([("Phone", FILLS["Duplicate"]), ("Name", ALT_FILLS["Duplicate"])]):
        r = 15 + i
        dc(ws2, r, 1, method, fill=fill, align="left")
        dc(ws2, r, 2, f'=COUNTIFS({LBL_R},"Duplicate",{METH_R},"{method}")', fill=fill)
        dc(ws2, r, 3, f'=IF(COUNTIF({LBL_R},"Duplicate")=0,0,B{r}/COUNTIF({LBL_R},"Duplicate"))', fill=fill, fmt="0.0%")

    sec(ws2, 19, 1, "DUPLICATES BY CRM STATUS", span=3)
    for ci, h in enumerate(["CRM Account Status", "Count", "% of Dupes"], 1):
        hdr(ws2, 20, ci, h)
    dup_statuses = df[df["Label"] == "Duplicate"]["CRM Account Status"].value_counts().index.tolist()
    for i, st in enumerate(dup_statuses):
        r    = 21 + i
        fill = ALT_FILLS["Duplicate"] if i % 2 == 0 else FILLS["Duplicate"]
        dc(ws2, r, 1, st, fill=fill, align="left")
        dc(ws2, r, 2, f'=COUNTIFS({LBL_R},"Duplicate",{CRMS_R},"{st}")', fill=fill)
        dc(ws2, r, 3, f'=IF(COUNTIF({LBL_R},"Duplicate")=0,0,B{r}/COUNTIF({LBL_R},"Duplicate"))', fill=fill, fmt="0.0%")

    sec(ws2, 29, 1, "TOP GM CATEGORIES", span=3)
    for ci, h in enumerate(["GM Category", "Count", "% of Matched"], 1):
        hdr(ws2, 30, ci, h)
    top_cats = df[df["GM Category"].notna() & (df["GM Category"] != "")]["GM Category"].value_counts().head(20).index.tolist()
    MATCHED_F = (
        f'=COUNTIF({LBL_R},"Qualified / Convert")+COUNTIF({LBL_R},"Duplicate")'
        f'+COUNTIF({LBL_R},"Business Closed")+COUNTIF({LBL_R},"Wrong Target Group")'
    )
    ws2.cell(row=52, column=2).value = MATCHED_F
    ws2.cell(row=52, column=2).font  = Font(color="FFFFFF", size=1)
    for i, cat in enumerate(top_cats):
        r    = 31 + i
        fill = PatternFill("solid", start_color="EBF3FB") if i % 2 == 0 else PatternFill("solid", start_color="FFFFFF")
        dc(ws2, r, 1, cat, fill=fill, align="left")
        dc(ws2, r, 2, f'=COUNTIF({CAT_R},"{cat}")', fill=fill)
        dc(ws2, r, 3, f"=IF(B52=0,0,B{r}/B52)", fill=fill, fmt="0.0%")

    sec(ws2, 4, 5, "TOP CITIES", span=6)
    for ci, h in enumerate(["City", "Qualified", "Duplicate", "Closed", "Wrong TG", "Total"], 5):
        hdr(ws2, 5, ci, h)
    for i, city in enumerate(df["City"].value_counts().head(20).index.tolist()):
        r    = 6 + i
        fill = PatternFill("solid", start_color="EBF3FB") if i % 2 == 0 else PatternFill("solid", start_color="FFFFFF")
        dc(ws2, r, 5, city, fill=fill, bold=True, align="left")
        dc(ws2, r, 6,  f'=COUNTIFS({LBL_R},"Qualified / Convert",{CITY_R},E{r})', fill=fill)
        dc(ws2, r, 7,  f'=COUNTIFS({LBL_R},"Duplicate",{CITY_R},E{r})', fill=fill)
        dc(ws2, r, 8,  f'=COUNTIFS({LBL_R},"Business Closed",{CITY_R},E{r})', fill=fill)
        dc(ws2, r, 9,  f'=COUNTIFS({LBL_R},"Wrong Target Group",{CITY_R},E{r})', fill=fill)
        dc(ws2, r, 10, f"=SUM(F{r}:I{r})", fill=fill, bold=True)

    for col, w in zip("ABCDE",  [28, 10, 12, 2, 22]):
        ws2.column_dimensions[col].width = w
    for col, w in zip("FGHIJ",  [14, 12, 12, 12, 10]):
        ws2.column_dimensions[col].width = w

    # ── Sheet 3: Qualified ───────────────────────────────────────
    ws3 = wb.create_sheet("✅ Qualified")
    q_df = df[df["Label"] == "Qualified / Convert"].reset_index(drop=True)
    ws3["A1"] = f"Qualified / Convert – Sales Ready  ({len(q_df):,} leads)"
    ws3["A1"].font = Font(name="Arial", bold=True, size=13, color="276221")
    ws3.merge_cells("A1:M1")
    q_heads = ["GRID", "Lead ID", "Company / Account", "City", "Street", "Phone",
               "GM Title", "GM Category", "GM Phone", "GM Website", "GM URL",
               "Match Confidence", "Match Reason"]
    for ci, h in enumerate(q_heads, 1):
        hdr(ws3, 3, ci, h)
    for ri, (_, row) in enumerate(q_df.iterrows(), 4):
        fill = FILLS["Qualified / Convert"] if ri % 2 == 0 else ALT_FILLS["Qualified / Convert"]
        for ci, key in enumerate(q_heads, 1):
            val = row.get(key, "")
            if pd.isna(val): val = ""
            c = ws3.cell(row=ri, column=ci, value=val if val != "" else "")
            c.border = thin()
            if key == "Match Confidence":
                try: c.value = float(val); c.number_format = "0.00"
                except Exception: pass
                c.font = Font(name="Arial", size=10, bold=True)
                c.fill = conf_fill(val)
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif key == "Match Reason":
                c.font = Font(name="Arial", size=9, color="595959")
                c.fill = fill; c.alignment = Alignment(vertical="center")
            else:
                c.font = Font(name="Arial", size=10)
                c.fill = fill; c.alignment = Alignment(vertical="center")
    for i, w in enumerate([10, 18, 32, 14, 36, 16, 30, 24, 16, 30, 50, 14, 38], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.freeze_panes = "A4"
    ws3.auto_filter.ref = f"A3:M{3 + len(q_df)}"

    # ── Sheet 4: Duplicates ──────────────────────────────────────
    ws4 = wb.create_sheet("🔴 Duplicates")
    d_df = df[df["Label"] == "Duplicate"].reset_index(drop=True)
    ws4["A1"] = f"Duplicates – Found in CRM  ({len(d_df):,} leads)"
    ws4["A1"].font = Font(name="Arial", bold=True, size=13, color="9C0006")
    ws4.merge_cells("A1:L1")
    d_heads = ["GRID", "Lead ID", "Company / Account", "City", "Phone",
               "Duplicate GRID", "Duplicate CRM Name", "CRM Account Status", "CRM Status Reason",
               "Duplicate Match Method", "GM Category", "GM Business Status"]
    for ci, h in enumerate(d_heads, 1):
        hdr(ws4, 3, ci, h)
    for ri, (_, row) in enumerate(d_df.iterrows(), 4):
        fill = FILLS["Duplicate"] if ri % 2 == 0 else ALT_FILLS["Duplicate"]
        for ci, key in enumerate(d_heads, 1):
            val = row.get(key, "")
            if pd.isna(val): val = ""
            c = ws4.cell(row=ri, column=ci, value=str(val) if val != "" else "")
            c.font = Font(name="Arial", size=10, color="9C0006")
            c.fill = fill; c.border = thin(); c.alignment = Alignment(vertical="center")
    for i, w in enumerate([10, 18, 32, 14, 16, 14, 30, 16, 18, 16, 24, 18], 1):
        ws4.column_dimensions[get_column_letter(i)].width = w
    ws4.freeze_panes = "A4"
    ws4.auto_filter.ref = f"A3:L{3 + len(d_df)}"

    # ── Sheet 5: Needs Review ────────────────────────────────────
    ws5 = wb.create_sheet("⚠️ Needs Review")
    rev_df = df[df["Label"].isin(["Business Closed", "Wrong Target Group", "Invalid Data"])].reset_index(drop=True)
    ws5["A1"] = f"Needs Review  ({len(rev_df):,} leads)"
    ws5["A1"].font = Font(name="Arial", bold=True, size=13, color="7D4E00")
    ws5.merge_cells("A1:M1")
    r_heads = ["GRID", "Lead ID", "Company / Account", "City", "Phone",
               "Label", "GM Title", "GM Category", "GM Business Status",
               "GM Phone", "Match Confidence", "Match Reason", "GM URL"]
    for ci, h in enumerate(r_heads, 1):
        hdr(ws5, 3, ci, h)
    for ri, (_, row) in enumerate(rev_df.iterrows(), 4):
        lbl  = row["Label"]
        fill = FILLS.get(lbl, PatternFill()) if ri % 2 == 0 else ALT_FILLS.get(lbl, PatternFill())
        lc   = LABEL_COLORS.get(lbl, "000000")
        for ci, key in enumerate(r_heads, 1):
            val = row.get(key, "")
            if pd.isna(val): val = ""
            c = ws5.cell(row=ri, column=ci, value=val if val != "" else "")
            c.border = thin()
            if key == "Label":
                c.font = Font(name="Arial", size=10, bold=True, color=lc)
                c.fill = fill; c.alignment = Alignment(horizontal="center", vertical="center")
            elif key == "Match Confidence":
                try: c.value = float(val); c.number_format = "0.00"
                except Exception: pass
                c.font = Font(name="Arial", size=10, bold=True)
                c.fill = conf_fill(val); c.alignment = Alignment(horizontal="center", vertical="center")
            elif key == "Match Reason":
                c.font = Font(name="Arial", size=9, color="595959")
                c.fill = fill; c.alignment = Alignment(vertical="center")
            else:
                c.font = Font(name="Arial", size=10)
                c.fill = fill; c.alignment = Alignment(vertical="center")
    for i, w in enumerate([10, 18, 32, 14, 16, 22, 30, 24, 18, 16, 14, 38, 50], 1):
        ws5.column_dimensions[get_column_letter(i)].width = w
    ws5.freeze_panes = "A4"
    ws5.auto_filter.ref = f"A3:M{3 + len(rev_df)}"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────
def main():
    st.title("🎯 Lead Classifier")
    st.caption("Upload your lead file, Apify results and CRM export — get a classified Excel report.")

    # ── Sidebar: Market selection ──────────────────────────────
    with st.sidebar:
        st.header("Settings")
        market_code = st.selectbox(
            "Market",
            options=list(MARKETS.keys()),
            format_func=lambda k: f"{MARKETS[k]['flag']} {MARKETS[k]['name']} ({k})",
        )
        market_cfg = MARKETS[market_code]
        st.caption(f"Phone prefix: +{market_cfg['phone_prefix']}  ·  Country: {market_cfg['country_suffix']}")

        st.divider()
        st.subheader("Match confidence threshold")
        confidence_threshold = st.slider(
            "Minimum name confidence score",
            min_value=0.3,
            max_value=1.0,
            value=0.5,
            step=0.05,
            help="Leads where the Google Maps name score is below this value are marked Invalid Data. Default 0.5 is recommended.",
        )
        if confidence_threshold < 0.5:
            st.warning(f"⚠️ Threshold set to {confidence_threshold:.2f} — matches below 0.5 may include incorrect businesses. Review the Match Confidence column carefully.")
        elif confidence_threshold >= 0.8:
            st.info(f"Strict mode ({confidence_threshold:.2f}) — only near-exact name matches will qualify.")
        else:
            st.success(f"Threshold: {confidence_threshold:.2f} (recommended)")

        st.divider()
        st.subheader("About")
        st.caption(
            "Classifies leads into:\n"
            "- ✅ Qualified / Convert\n"
            "- 🔴 Duplicate\n"
            "- 🟡 Business Closed\n"
            "- 🟠 Wrong Target Group\n"
            "- ⚫ Invalid Data"
        )

    # ── Tabs ──────────────────────────────────────────────────
    tab_classify, tab_urls = st.tabs(["📊 Classify leads", "🔗 Generate Apify URLs"])

    # ── TAB 1: Classify ───────────────────────────────────────
    with tab_classify:

        # ── Help links ────────────────────────────────────────
        with st.expander("📎 How to get your files — click to expand"):
            c1, c2, c3, c4 = st.columns(4)

            with c1:
                st.markdown("**1. CRM export (Salesforce)**")
                st.markdown(
                    "Open the report, change the **Account Country** filter to your market, "
                    "then export as CSV. For more than 100k rows use Salesforce Inspector (see step 4)."
                )
                st.link_button(
                    "Open CRM Report →",
                    "https://deliveryhero.lightning.force.com/lightning/r/Report/00ObO0000047MEPUA2/view?queryScope=userFolders",
                    use_container_width=True,
                )

            with c2:
                st.markdown("**2. Leads export (Salesforce)**")
                st.markdown(
                    "Open the report, change the **Country** and **Lead Source** filters to match your market, "
                    "then export as CSV."
                )
                st.link_button(
                    "Open Leads Report →",
                    "https://deliveryhero.lightning.force.com/lightning/r/Report/00ObO0000047LqDUAU/view?queryScope=userFolders",
                    use_container_width=True,
                )

            with c3:
                st.markdown("**3. Apify — Google Maps Extractor**")
                st.markdown(
                    "Use the **URL generator tab** above to create your search URLs first. "
                    "Paste them into Apify and run. When exporting the results, make sure "
                    "these fields are selected:"
                )
                st.code(
                    "title, temporarilyClosed, permanentlyClosed,\n"
                    "postalCode, address, city, street,\n"
                    "website, phone, phoneUnformatted,\n"
                    "categoryName, categories, url, searchPageUrl",
                    language="text",
                )
                st.caption("Format: CSV · View: Overview · additionalInfo can be omitted to keep file size small.")
                st.link_button(
                    "Open Apify Actor →",
                    "https://console.apify.com/actors/compass~crawler-google-places",
                    use_container_width=True,
                )

            with c4:
                st.markdown("**4. CRM export > 100k rows — Salesforce Inspector**")
                st.markdown(
                    "The Salesforce report is capped at 100k rows. "
                    "For a full export install the Chrome extension below, "
                    "open it on any Salesforce page, go to the **Export** tab and run:"
                )
                st.code(
                    "SELECT GRID__c, Name, Phone,\n"
                    "Account_Status__c, Status_Reason__c,\n"
                    "BillingCity\n"
                    "FROM Account\n"
                    "WHERE BillingCountry = 'YOUR_COUNTRY'",
                    language="sql",
                )
                st.caption("Replace YOUR_COUNTRY with e.g. Norway, Turkey, Sweden etc. Download as CSV when done.")
                st.link_button(
                    "Salesforce Inspector Chrome Extension →",
                    "https://chromewebstore.google.com/detail/salesforce-inspector-reloaded/hpijlohoihegkfehhibggnkbjhoemldh",
                    use_container_width=True,
                )

        st.divider()
        col1, col2, col3 = st.columns(3)

        with col1:
            st.subheader("1. ALG Leads")
            leads_file = st.file_uploader("Upload leads file (.xlsx or .csv)", type=["xlsx", "csv"], key="leads")

        with col2:
            st.subheader("2. Apify Results")
            apify_file = st.file_uploader("Upload Apify Google Maps output (.csv or .xlsx)", type=["csv", "xlsx"], key="apify")
            st.caption("Optional — skipped if not uploaded.")

        with col3:
            st.subheader("3. CRM Export")
            crm_file = st.file_uploader("Upload Salesforce CRM export (.csv or .xlsx)", type=["csv", "xlsx"], key="crm")
            st.caption("Optional — duplicate check skipped if not uploaded.")

        if leads_file:
            try:
                leads_df, col_map_leads = load_leads(leads_file, market_cfg)
                st.success(f"Leads loaded: {len(leads_df):,} rows")

                # Show detected columns
                with st.expander("Detected column mapping — click to review"):
                    col_df = pd.DataFrame([
                        {"Field": k, "Detected column": v or "⚠️ Not found"}
                        for k, v in col_map_leads.items()
                    ])
                    st.dataframe(col_df, hide_index=True, use_container_width=True)
            except Exception as e:
                st.error(f"Error loading leads file: {e}")
                leads_df = None
                col_map_leads = {}
        else:
            leads_df = None
            col_map_leads = {}

        crm_df, col_map_crm = None, {}
        if crm_file:
            try:
                crm_df, col_map_crm = load_crm(crm_file, market_cfg)
                st.success(f"CRM loaded: {len(crm_df):,} accounts")
            except Exception as e:
                st.error(f"Error loading CRM file: {e}")

        apify_df, col_map_apify = None, {}
        if apify_file:
            try:
                apify_df, col_map_apify = load_apify(apify_file)
                st.success(f"Apify results loaded: {len(apify_df):,} rows")
            except Exception as e:
                st.error(f"Error loading Apify file: {e}")

        st.divider()

        if leads_df is not None and st.button("▶ Run classification", type="primary", use_container_width=True):
            with st.spinner("Classifying leads..."):
                result_df = classify_leads(
                    leads_df, col_map_leads,
                    crm_df, col_map_crm,
                    apify_df, col_map_apify,
                    market_cfg,
                    confidence_threshold=confidence_threshold,
                )

            st.success("Done!")
            counts = result_df["Label"].value_counts()

            # ── Summary metrics ────────────────────────────────
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("✅ Qualified",      counts.get("Qualified / Convert", 0))
            c2.metric("🔴 Duplicate",      counts.get("Duplicate", 0))
            c3.metric("🟡 Business Closed",counts.get("Business Closed", 0))
            c4.metric("🟠 Wrong TG",       counts.get("Wrong Target Group", 0))
            c5.metric("⚫ Invalid Data",   counts.get("Invalid Data", 0))

            if crm_df is not None:
                dup_df = result_df[result_df["Label"] == "Duplicate"]
                st.caption(
                    f"Duplicates: {dup_df['Duplicate Match Method'].value_counts().get('Phone', 0)} by phone  ·  "
                    f"{dup_df['Duplicate Match Method'].value_counts().get('Name', 0)} by exact name"
                )

            # ── Download button ────────────────────────────────
            excel_buf = build_excel(result_df, f"{market_cfg['flag']} {market_cfg['name']}")
            st.download_button(
                label="⬇ Download Excel report",
                data=excel_buf,
                file_name=f"leads_classified_{market_code}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            # ── Preview ────────────────────────────────────────
            with st.expander("Preview results"):
                st.dataframe(
                    result_df[["Company / Account", "City", "Label", "GM Category",
                                "GM Business Status", "Match Confidence", "Duplicate GRID"]],
                    hide_index=True,
                    use_container_width=True,
                )

    # ── TAB 2: URL Generator ──────────────────────────────────
    with tab_urls:
        st.subheader("Generate Google Maps URLs for Apify")
        st.caption(
            "Upload your leads file to generate search URLs. "
            "Paste them into Apify's Google Maps Extractor, run it, then come back to Tab 1."
        )

        url_leads_file = st.file_uploader(
            "Upload leads file (.xlsx or .csv)", type=["xlsx", "csv"], key="url_leads"
        )

        if url_leads_file:
            try:
                url_leads_df, url_col_map = load_leads(url_leads_file, market_cfg)
                urls = generate_google_urls(url_leads_df, url_col_map, market_cfg)
                url_leads_df["GOOGLE URL"] = urls

                method = "coordinates" if url_col_map.get("lat") else "name + address"
                st.success(f"{len(urls):,} URLs generated using {method}.")

                # Preview
                preview = url_leads_df[[
                    c for c in [url_col_map.get("name"), url_col_map.get("city"), "GOOGLE URL"]
                    if c
                ]].head(10)
                st.dataframe(preview, hide_index=True, use_container_width=True)

                # Download as CSV
                csv_buf = io.StringIO()
                url_leads_df.to_csv(csv_buf, index=False)
                st.download_button(
                    label="⬇ Download leads with Google URLs (.csv)",
                    data=csv_buf.getvalue(),
                    file_name=f"leads_with_urls_{market_code}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

                st.info(
                    "**Next steps:**\n"
                    "1. Download the CSV above\n"
                    "2. Open Apify → Google Maps Extractor\n"
                    "3. Paste the URLs from the GOOGLE URL column as input\n"
                    "4. Run the actor and download the results CSV\n"
                    "5. Come back to the **Classify leads** tab and upload it"
                )
            except Exception as e:
                st.error(f"Error: {e}")


if __name__ == "__main__":
    main()
