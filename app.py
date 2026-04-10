import streamlit as st
import pandas as pd
import re
import io
import time
import json
from urllib.parse import unquote, quote, urlencode
from urllib.request import urlopen
from difflib import SequenceMatcher
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from shapely.wkt import loads as wkt_loads
from shapely.geometry import Point

# ── App config ──────────────────────────────────────────────────
APP_PASSWORD = "Pandora2026"
def _load_logo():
    """Load DH logo from file, return base64 string for embedding."""
    import os, base64
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dh_logo.png")
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""

DH_LOGO_B64 = _load_logo()


def check_password():
    """Show password gate. Returns True if correct password entered."""
    if st.session_state.get("authenticated"):
        return True

    # Centre the login card
    _, col, _ = st.columns([1, 1.6, 1])
    with col:
        st.markdown(f'''
        <div style="text-align:center; padding: 2rem 0 1rem 0;">
            <img src="data:image/png;base64,{DH_LOGO_B64}" style="width:160px; margin-bottom:1.5rem;" />
            <h2 style="color:#1A1A1A; font-size:1.4rem; font-weight:700; margin-bottom:0.3rem;">Lead Classifier</h2>
            <p style="color:#888; font-size:0.9rem; margin-bottom:1.5rem;">Digital Sales EU · Pandora / Delivery Hero</p>
        </div>''', unsafe_allow_html=True)

        pwd = st.text_input("Password", type="password", placeholder="Enter password to continue")
        if st.button("Sign in", type="primary", use_container_width=True):
            if pwd == APP_PASSWORD:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Incorrect password — please try again.")
    return False


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
        "code": "TR",
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
        "code": "NO",
        "name": "Norway",
        "flag": "🇳🇴",
        "char_map": {
            "Æ": "AE", "Ø": "O", "Å": "A", "Ü": "U", "Ö": "O", "Ä": "A",
            "æ": "ae", "ø": "o", "å": "a", "ü": "u", "ö": "o", "ä": "a",
        },
        "country_suffix": "Norway",
        "phone_prefix": "47",
    },
    "SE": {
        "code": "SE",
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
        "code": "CZ",
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
        "code": "HU",
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
        "code": "AT",
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

# Allowlist approach — only these categories are considered Foodora-eligible.
# Anything not in this set is marked Wrong Target Group.
FOOD_DELIVERY_ALLOWED = {
    # Restaurants — general
    "Restaurant", "Fine dining restaurant", "Family restaurant", "Casual dining restaurant",
    "Buffet restaurant", "Brasserie", "Bistro", "Diner", "Eatery",

    # Restaurants — cuisine specific
    "Turkish restaurant", "Italian restaurant", "French restaurant", "Chinese restaurant",
    "Japanese restaurant", "Thai restaurant", "Indian restaurant", "Mexican restaurant",
    "Greek restaurant", "Lebanese restaurant", "Middle Eastern restaurant",
    "Mediterranean restaurant", "Asian restaurant", "Korean restaurant",
    "Vietnamese restaurant", "Spanish restaurant", "American restaurant",
    "Ethiopian restaurant", "Afghan restaurant", "Pakistani restaurant",
    "Nepalese restaurant", "Sri Lankan restaurant", "Bangladeshi restaurant",
    "Indonesian restaurant", "Filipino restaurant", "Peruvian restaurant",
    "Brazilian restaurant", "Argentinian restaurant", "Georgian restaurant",
    "Uzbek restaurant", "Syrian restaurant", "Iraqi restaurant",
    "Moroccan restaurant", "Egyptian restaurant", "Tunisian restaurant",
    "Algerian restaurant", "Libyan restaurant", "Yemeni restaurant",
    "Somali restaurant", "Nigerian restaurant", "Ghanaian restaurant",
    "Senegalese restaurant", "Cameroonian restaurant",
    "Caribbean restaurant", "Jamaican restaurant", "Cuban restaurant",
    "Portuguese restaurant", "German restaurant", "Austrian restaurant",
    "Czech restaurant", "Hungarian restaurant", "Polish restaurant",
    "Romanian restaurant", "Bulgarian restaurant", "Croatian restaurant",
    "Serbian restaurant", "Bosnian restaurant", "Macedonian restaurant",
    "Albanian restaurant", "Russian restaurant", "Ukrainian restaurant",
    "Scandinavian restaurant", "Nordic restaurant",
    "Latin American restaurant", "Fusion restaurant", "International restaurant",
    "European restaurant", "Pan-Asian restaurant", "Oriental restaurant",

    # Fast food & quick service
    "Fast food restaurant", "Fast-food restaurant", "Quick service restaurant",
    "Hamburger restaurant", "Burger restaurant", "Hot dog restaurant",
    "Fried chicken restaurant", "Chicken restaurant", "Chicken wings restaurant",
    "Fish and chips restaurant", "Seafood restaurant", "Fish restaurant",
    "Taco restaurant", "Burrito restaurant",

    # Specific food types
    "Pizza restaurant", "Pizza delivery", "Pizza takeaway",
    "Kebab shop", "Kebab restaurant", "Doner kebab restaurant",
    "Shawarma restaurant", "Falafel restaurant", "Pita restaurant",
    "Sushi restaurant", "Ramen restaurant", "Noodle restaurant",
    "Dumpling restaurant", "Dim sum restaurant", "Wonton restaurant",
    "Steak house", "Steakhouse", "Grill restaurant", "Barbecue restaurant",
    "BBQ restaurant", "Smokehouse", "Rotisserie chicken restaurant",
    "Sandwich shop", "Submarine sandwich shop", "Wrap restaurant",
    "Salad shop", "Bowl restaurant", "Poke bar",
    "Soup restaurant", "Soup kitchen",
    "Breakfast restaurant", "Brunch restaurant",
    "Pancake restaurant", "Waffle restaurant",
    "Dessert restaurant", "Dessert shop",
    "Ice cream shop", "Ice cream parlor", "Frozen yogurt shop",
    "Donut shop", "Doughnut shop",
    "Crepe restaurant", "Waffle house",
    "Vegetarian restaurant", "Vegan restaurant", "Plant-based restaurant",
    "Organic restaurant", "Health food restaurant", "Gluten-free restaurant",
    "Halal restaurant", "Kosher restaurant",

    # Cafe & coffee
    "Cafe", "Coffee shop", "Coffee house", "Coffeehouse",
    "Espresso bar", "Tea house", "Bubble tea shop", "Boba shop",
    "Internet cafe", "Café",

    # Bakery & pastry
    "Bakery", "Patisserie", "Pastry shop", "Cake shop",
    "Bread bakery", "Artisan bakery", "French bakery",
    "Donut shop", "Cookie shop", "Cupcake shop",
    "Bagel shop",

    # Delivery & takeaway oriented
    "Meal delivery", "Food delivery", "Takeaway",
    "Takeout restaurant", "Take-out restaurant",
    "Catering", "Caterer", "Catering food and drink supplier",
    "Cloud kitchen", "Ghost kitchen", "Virtual restaurant",

    # Bars with food (eligible if they serve food)
    "Bar & grill", "Bar and grill", "Sports bar", "Gastropub",
    "Pub", "Tavern", "Tapas bar", "Wine bar",

    # Specific local formats
    "Köfte restaurant", "Kofta restaurant", "Lahmacun restaurant",
    "Pide restaurant", "Börek shop", "Gözleme restaurant",
    "Manti restaurant", "Tantuni restaurant", "Dürüm restaurant",
    "Iskender restaurant", "Adana kebab restaurant",
    "Tripe restaurant", "Offal restaurant",
    "Mussel restaurant", "Shrimp restaurant",
    "Kokoreç", "Simit shop",
    "Smørbrød restaurant", "Smørrebrød restaurant",
    "Wiener schnitzel restaurant", "Schnitzel restaurant",
    "Goulash restaurant", "Langos", "Kürtőskalács",
    "Trdelník shop", "Svíčková restaurant",
    "Pierogi restaurant", "Żurek restaurant",
    "Knedlíky restaurant", "Svíčková",

    # General / catch-all food formats
    "Food court", "Food hall", "Food stall", "Food truck",
    "Street food restaurant", "Market restaurant",
    "Home cooking restaurant", "Traditional restaurant",
    "Local restaurant", "Neighborhood restaurant",
    "Deli", "Delicatessen", "Charcuterie",
    "Canteen", "Staff canteen", "University canteen",
    "School cafeteria", "Hospital cafeteria", "Cafeteria",
    "Lunchroom", "Snack bar",
    "Juice bar", "Smoothie bar", "Açaí shop",

    # Specialty
    "Chocolate shop", "Sweet shop", "Candy store",
    "Noodle shop", "Pasta shop", "Rice restaurant",
    "Porridge restaurant", "Congee restaurant",
    "Hot pot restaurant", "Fondue restaurant", "Raclette restaurant",
    "Teppanyaki restaurant", "Okonomiyaki restaurant", "Takoyaki restaurant",
    "Yakitori restaurant", "Izakaya", "Robatayaki restaurant",
    "Tempura restaurant", "Tonkatsu restaurant", "Udon restaurant",
    "Soba restaurant", "Gyoza restaurant",
    "Pho restaurant", "Banh mi restaurant", "Spring roll restaurant",
    "Satay restaurant", "Rendang restaurant",
    "Curry restaurant", "Tandoori restaurant", "Biryani restaurant",
    "Dosa restaurant", "Idli restaurant", "Thali restaurant",
    "Ceviche restaurant", "Empanada restaurant",
    "Arepas restaurant", "Chimichanga restaurant",
}

def is_food_delivery_eligible(category):
    """Return True if the category is on the Foodora-eligible allowlist."""
    if not category or pd.isna(category):
        return False
    cat = str(category).strip()
    # Exact match
    if cat in FOOD_DELIVERY_ALLOWED:
        return True
    # Case-insensitive fallback
    cat_lower = cat.lower()
    for allowed in FOOD_DELIVERY_ALLOWED:
        if allowed.lower() == cat_lower:
            return True
    # Contains key food words — catches long compound category names
    food_keywords = [
        "restaurant", "cafe", "café", "bakery", "kebab", "pizza", "sushi",
        "burger", "grill", "bistro", "brasserie", "diner", "eatery", "kitchen",
        "takeaway", "takeout", "delivery", "catering", "patisserie", "pastry",
        "coffee", "tea house", "noodle", "ramen", "deli", "canteen", "cafeteria",
        "snack", "food truck", "street food", "bar & grill", "gastropub",
        "steakhouse", "steak house", "seafood", "sandwich", "pide", "lahmacun",
        "köfte", "döner", "shawarma", "falafel", "taco", "burrito",
    ]
    for kw in food_keywords:
        if kw in cat_lower:
            return True
    return False

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
    # Reject dummy/placeholder numbers before adding prefix
    digits_raw = re.sub(r"\D", "", s)
    if len(digits_raw) < 6:              return ""   # too short
    if len(set(digits_raw)) == 1:        return ""   # all same digit: 00000, 11111
    if digits_raw.endswith("0" * 6):     return ""   # ends in 6+ zeros: +4700000000
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


# Generic words that appear in many business names — strip before token matching
# so "Pizzeria X" vs "Pizzeria Y" doesn't score high just because of shared generic words
_GENERIC_NAME_TOKENS = {
    "restaurant", "cafe", "pizza", "kebab", "kebap", "bar", "grill", "ristorante",
    "bistro", "house", "shop", "the", "and", "di", "da", "de", "le", "la", "al",
    "asian", "chinese", "italian", "mexican", "thai", "indian", "pizzeria",
    "gasthaus", "gasthof", "wirtshaus", "kafe", "lokanta", "restoran",
}


def name_confidence(a, b, char_map):
    a = norm_name(a, char_map)
    b = norm_name(b, char_map)
    if not a or not b:
        return 0.0

    raw_a = set(re.split(r"\W+", a)) - {""}
    raw_b = set(re.split(r"\W+", b)) - {""}

    # Strip generic tokens for token matching, fall back to full set if nothing remains
    tokens_a = raw_a - _GENERIC_NAME_TOKENS or raw_a
    tokens_b = raw_b - _GENERIC_NAME_TOKENS or raw_b

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


def find_header_row(file, key_col="GRID"):
    """Scan first 30 rows to find where key_col header appears — handles Salesforce report exports."""
    try:
        raw = pd.read_excel(file, header=None, nrows=30)
        for i, row in raw.iterrows():
            if any(str(v).strip() == key_col for v in row.dropna()):
                return i
    except Exception:
        pass
    return 0


def load_leads(file, market_cfg):
    """Load leads file and auto-detect columns."""
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        header_row = find_header_row(file, key_col="GRID")
        df = pd.read_excel(file, header=header_row)
        # Drop Salesforce footer rows — keep only rows where GRID looks like a real ID
        grid_col = next((c for c in df.columns if str(c).strip() == "GRID"), None)
        if grid_col:
            df = df[df[grid_col].astype(str).str.match(r'^[A-Z0-9]{6,}$')].copy()

    col_map = {}
    col_map["name"]    = detect_column(df, ["Company / Account", "Account Name", "Name", "company_name", "Företag", "Virksomhed", "Vállalkozás", "Unternehmen"])
    col_map["phone"]   = detect_column(df, ["Phone", "phone_number", "Telefon", "Telefonnummer", "Mobile"])
    col_map["street"]  = detect_column(df, ["Street", "Address", "address", "Adresse", "Cím"])
    col_map["city"]    = detect_column(df, ["City", "city", "By", "Stad", "Město", "Város", "Stadt", "Restaurant City"])
    col_map["grid"]    = detect_column(df, ["GRID", "grid", "Grid"])
    col_map["lead_id"] = detect_column(df, ["Lead ID", "lead_id", "LeadID", "ID"])
    col_map["url"]     = detect_column(df, ["GOOGLE URL", "Google URL", "google_url", "URL"])
    col_map["lat"]     = detect_column(df, ["Coordinates (Latitude)", "Latitude", "lat"])
    col_map["lng"]     = detect_column(df, ["Coordinates (Longitude)", "Longitude", "lng"])
    col_map["zip"]     = detect_column(df, ["Zip/Postal Code", "Zip", "postal_code", "PostalCode", "Postnummer", "PSČ", "Irányítószám", "PLZ", "Postinummer"])

    return df, col_map


def load_crm(file, market_cfg):
    """Load CRM file, handling Salesforce report headers."""
    prefix = market_cfg["phone_prefix"]
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        header_row = find_header_row(file, key_col="GRID")
        df = pd.read_excel(file, header=header_row)
        # Drop Salesforce footer rows — keep only rows where GRID looks like a real ID
        grid_col = next((c for c in df.columns if str(c).strip() == "GRID"), None)
        if grid_col:
            df = df[df[grid_col].astype(str).str.match(r'^[A-Z0-9]{6,}$')].copy()

    col_map = {}
    col_map["grid"]   = detect_column(df, ["GRID__c", "GRID", "Grid"])
    col_map["name"]   = detect_column(df, ["Account Name", "Name", "name"])
    col_map["phone"]  = detect_column(df, ["Phone", "phone"])
    col_map["status"] = detect_column(df, ["Account_Status__c", "Account Status", "AccountStatus"])
    col_map["reason"] = detect_column(df, ["Status_Reason__c", "Status Reason", "StatusReason"])
    col_map["city"]   = detect_column(df, ["BillingCity", "Restaurant City", "City", "city", "Stad", "By", "Město", "Város", "Stadt"])

    phone_col = col_map["phone"]
    if phone_col:
        df["_phone_norm"] = df[phone_col].apply(lambda p: norm_phone(p, prefix))

    name_col = col_map["name"]
    char_map = market_cfg["char_map"]
    if name_col:
        df["_name_norm"] = df[name_col].apply(lambda n: norm_name(n, char_map))

    city_col = col_map["city"]
    if city_col:
        df["_city_norm"] = df[city_col].apply(lambda c: norm_city(c, char_map))
    else:
        df["_city_norm"] = ""

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


# District → Province mappings per market
# Handles cases where leads use district names but CRM stores province
_DISTRICT_MAP = {
    "TR": {
        "BESIKTAS":"ISTANBUL","KADIKOY":"ISTANBUL","USKUDAR":"ISTANBUL","BEYOGLU":"ISTANBUL",
        "SISLI":"ISTANBUL","BAKIRKOY":"ISTANBUL","FATIH":"ISTANBUL","EYUP":"ISTANBUL",
        "EYUPSULTAN":"ISTANBUL","GAZIOSMANPASA":"ISTANBUL","BAGCILAR":"ISTANBUL",
        "BAHCELIEVLER":"ISTANBUL","BAYRAMPASA":"ISTANBUL","ESENLER":"ISTANBUL",
        "GUNGOREN":"ISTANBUL","KUCUKCEKMECE":"ISTANBUL","AVCILAR":"ISTANBUL",
        "BUYUKCEKMECE":"ISTANBUL","CATALCA":"ISTANBUL","ESENYURT":"ISTANBUL",
        "ARNAVUTKOY":"ISTANBUL","BASAKSEHIR":"ISTANBUL","BEYLIKDUZU":"ISTANBUL",
        "SULTANGAZI":"ISTANBUL","ZEYTINBURNU":"ISTANBUL","SARIYER":"ISTANBUL",
        "BEYKOZ":"ISTANBUL","ADALAR":"ISTANBUL","ATASEHIR":"ISTANBUL",
        "KARTAL":"ISTANBUL","MALTEPE":"ISTANBUL","PENDIK":"ISTANBUL",
        "SULTANBEYLI":"ISTANBUL","TUZLA":"ISTANBUL","CEKMEKOY":"ISTANBUL",
        "SANCAKTEPE":"ISTANBUL","UMRANIYE":"ISTANBUL",
        "CANKAYA":"ANKARA","KECIOREN":"ANKARA","MAMAK":"ANKARA","YENIMAHALLE":"ANKARA",
        "ETIMESGUT":"ANKARA","SINCAN":"ANKARA","ALTINDAG":"ANKARA","PURSAKLAR":"ANKARA",
        "GOLBASI":"ANKARA","AKYURT":"ANKARA","KAZAN":"ANKARA","CUBUK":"ANKARA",
        "BORNOVA":"IZMIR","BUCA":"IZMIR","KARSIYAKA":"IZMIR","KONAK":"IZMIR",
        "BAYRAKLI":"IZMIR","GAZIEMIR":"IZMIR","CIGLI":"IZMIR","MENEMEN":"IZMIR",
        "BERGAMA":"IZMIR","EFELER":"IZMIR","ALSANCAK":"IZMIR",
        "OSMANGAZI":"BURSA","NILUFER":"BURSA","YILDIRIM":"BURSA","INEGOL":"BURSA",
        "GEMLIK":"BURSA","MUDANYA":"BURSA",
        "IZMIT":"KOCAELI","GEBZE":"KOCAELI","DERINCE":"KOCAELI","BASISKELE":"KOCAELI","KARTEPE":"KOCAELI",
        "MURATPASA":"ANTALYA","KEPEZ":"ANTALYA","KONYAALTI":"ANTALYA","ALANYA":"ANTALYA",
        "MANAVGAT":"ANTALYA","SERIK":"ANTALYA","KAS":"ANTALYA","KEMER":"ANTALYA",
        "SEYHAN":"ADANA","YUREGIR":"ADANA","CUKUROVA":"ADANA","KOZAN":"ADANA",
        "SAHINBEY":"GAZIANTEP","SEHITKAMIL":"GAZIANTEP",
        "MEZITLI":"MERSIN","YENISEHIR":"MERSIN","TOROSLAR":"MERSIN","AKDENIZ":"MERSIN",
        "ERDEMLI":"MERSIN","TARSUS":"MERSIN",
        "ONIKISUBAT":"KAHRAMANMARAS","DULKADIROGLU":"KAHRAMANMARAS",
        "ATAKUM":"SAMSUN","ILKADIM":"SAMSUN","CANIK":"SAMSUN","BAFRA":"SAMSUN",
        "KAYAPINAR":"DIYARBAKIR","BAGLAR":"DIYARBAKIR",
        "BODRUM":"MUGLA","MILAS":"MUGLA","MARMARIS":"MUGLA",
        "ANTAKYA":"HATAY","DORTYOL":"HATAY",
        "HENDEK":"SAKARYA","ERENLER":"SAKARYA","KOCAALI":"SAKARYA",
        "FATSA":"ORDU","ARDESEN":"RIZE","TATVAN":"BITLIS",
        "ERCIS":"VAN","CUMRA":"KONYA","BEYSEHIR":"KONYA",
        "BOLVADIN":"AFYONKARAHISAR","EDREMIT":"BALIKESIR","SOMA":"MANISA",
        "ERBAA":"TOKAT","NIKSAR":"TOKAT","VEZIRKOPRU":"SAMSUN",
        "SORGUN":"YOZGAT","ELBISTAN":"KAHRAMANMARAS","GONEN":"BALIKESIR",
    },
    "NO": {
        # Oslo districts
        "FROGNER":"OSLO","GRUNERLOKKA":"OSLO","GRUNERLØKKA":"OSLO","SAGENE":"OSLO",
        "NORDRE AKER":"OSLO","ST HANSHAUGEN":"OSLO","GAMLE OSLO":"OSLO",
        "GRONLAND":"OSLO","GRØNLAND":"OSLO","MAJORSTUEN":"OSLO","SENTRUM":"OSLO",
        "HOLMLIA":"OSLO","ULLERN":"OSLO","VESTRE AKER":"OSLO",
        "OSTENSJØ":"OSLO","ØSTENSJØ":"OSLO","NORDSTRAND":"OSLO",
        "SØNDRE NORDSTRAND":"OSLO","SONDRE NORDSTRAND":"OSLO",
        "BJERKE":"OSLO","GRORUD":"OSLO","STOVNER":"OSLO","ALNA":"OSLO",
        # Greater Oslo
        "LØRENSKOG":"LILLESTROM","LORENSKOG":"LILLESTROM",
        "SKEDSMO":"LILLESTROM","RÆLINGEN":"LILLESTROM","RALINGEN":"LILLESTROM",
        # Bergen districts
        "BERGENHUS":"BERGEN","YTREBYGDA":"BERGEN","FANA":"BERGEN",
        "LAKSEVAG":"BERGEN","LAKSEVAAG":"BERGEN",
        "ASANE":"BERGEN","ÅSANE":"BERGEN","ARNA":"BERGEN","FYLLINGSDALEN":"BERGEN",
        # Trondheim
        "MIDTBYEN":"TRONDHEIM","BYASEN":"TRONDHEIM","HEIMDAL":"TRONDHEIM",
        # Stavanger
        "HILLEVAG":"STAVANGER","EIGANES":"STAVANGER","HINNA":"STAVANGER",
        "TASTA":"STAVANGER","STORHAUG":"STAVANGER",
    },
    "SE": {
        # Stockholm districts
        "SODERMALM":"STOCKHOLM","SÖDERMALM":"STOCKHOLM",
        "VASASTAN":"STOCKHOLM","OSTERMALM":"STOCKHOLM","ÖSTERMALM":"STOCKHOLM",
        "KUNGSHOLMEN":"STOCKHOLM","NORRMALM":"STOCKHOLM","GAMLA STAN":"STOCKHOLM",
        "BROMMA":"STOCKHOLM","HAGERSTEN":"STOCKHOLM","HÄGERSTEN":"STOCKHOLM",
        "SKARPNACK":"STOCKHOLM","SKARPNÄCK":"STOCKHOLM","FARSTA":"STOCKHOLM",
        "ENSKEDE":"STOCKHOLM","ARSTA":"STOCKHOLM","ÅRSTA":"STOCKHOLM",
        "VANTOR":"STOCKHOLM","VANTÖR":"STOCKHOLM",
        "SPANGA":"STOCKHOLM","SPÅNGA":"STOCKHOLM","TENSTA":"STOCKHOLM",
        "RINKEBY":"STOCKHOLM","KISTA":"STOCKHOLM","HUSBY":"STOCKHOLM",
        # Swedish county names → main city
        "STOCKHOLMS LAN":"STOCKHOLM","STOCKHOLMS LÄN":"STOCKHOLM",
        "VASTRA GOTALANDS LAN":"GOTEBORG","VÄSTRA GÖTALANDS LÄN":"GOTEBORG",
        "SKANE LAN":"MALMO","SKÅNE LÄN":"MALMO",
        "OSTERGOTLANDS LAN":"LINKOPING","ÖSTERGÖTLANDS LÄN":"LINKOPING",
        "OREBRO LAN":"OREBRO","ÖREBRO LÄN":"OREBRO",
        "VASTMANLANDS LAN":"VASTERAS","VÄSTMANLANDS LÄN":"VASTERAS",
        "GAVLEBORGS LAN":"GAVLE","GÄVLEBORGS LÄN":"GAVLE",
        "JAMTLANDS LAN":"OSTERSUND","JÄMTLANDS LÄN":"OSTERSUND",
        "VASTERNORRLANDS LAN":"SUNDSVALL","VÄSTERNORRLANDS LÄN":"SUNDSVALL",
        "VASTERBOTTENS LAN":"UMEA","VÄSTERBOTTENS LÄN":"UMEA",
        "NORRBOTTENS LAN":"LULEA","NORRBOTTENS LÄN":"LULEA",
        "JONKOPINGS LAN":"JONKOPING","JÖNKÖPINGS LÄN":"JONKOPING",
        "KRONOBERGS LAN":"VAXJO","KRONOBERGS LÄN":"VAXJO",
        "KALMAR LAN":"KALMAR","KALMAR LÄN":"KALMAR",
        "GOTLANDS LAN":"VISBY","GOTLANDS LÄN":"VISBY",
        "BLEKINGE LAN":"KARLSKRONA","BLEKINGE LÄN":"KARLSKRONA",
        "HALLANDS LAN":"HALMSTAD","HALLANDS LÄN":"HALMSTAD",
        "VARMLANDS LAN":"KARLSTAD","VÄRMLANDS LÄN":"KARLSTAD",
        "DALARNAS LAN":"FALUN","DALARNAS LÄN":"FALUN",
        "SODERMANLANDS LAN":"ESKILSTUNA","SÖDERMANLANDS LÄN":"ESKILSTUNA",
        # Greater Stockholm municipalities (these appear as their own cities in CRM)
        "JARFALLA":"JARFALLA","JÄRFÄLLA":"JARFALLA",
        "LIDINGO":"LIDINGO","LIDINGÖ":"LIDINGO",
        "SOLLENTUNA":"SOLLENTUNA","UPPLANDS VASBY":"UPPLANDS VASBY",
        "UPPLANDS VÄSBY":"UPPLANDS VASBY",
        # Göteborg districts
        "HISINGEN":"GOTEBORG","MAJORNA":"GOTEBORG","CENTRUM":"GOTEBORG",
        "ASKIM":"GOTEBORG","FROLUNDA":"GOTEBORG","FRÖLUNDA":"GOTEBORG",
        "ANGERED":"GOTEBORG",
    },
    "AT": {
        # Vienna districts (Bezirke)
        "INNERE STADT":"WIEN","LEOPOLDSTADT":"WIEN","LANDSTRASSE":"WIEN",
        "WIEDEN":"WIEN","MARGARETEN":"WIEN","MARIAHILF":"WIEN",
        "NEUBAU":"WIEN","JOSEFSTADT":"WIEN","ALSERGRUND":"WIEN",
        "FAVORITEN":"WIEN","SIMMERING":"WIEN","MEIDLING":"WIEN",
        "HIETZING":"WIEN","PENZING":"WIEN","RUDOLFSHEIM":"WIEN",
        "OTTAKRING":"WIEN","HERNALS":"WIEN","WAEHRING":"WIEN","WÄHRING":"WIEN",
        "DOBLING":"WIEN","DÖBLING":"WIEN","BRIGITTENAU":"WIEN",
        "FLORIDSDORF":"WIEN","DONAUSTADT":"WIEN","LIESING":"WIEN",
        # Graz districts
        "INNENSTADT":"GRAZ","GRIES":"GRAZ","LEND":"GRAZ",
        "JAKOMINI":"GRAZ","EGGENBERG":"GRAZ","WETZELSDORF":"GRAZ",
        "LIEBENAU":"GRAZ","PUNTIGAM":"GRAZ","MARIATROST":"GRAZ",
        # Long Austrian city names — normalise to consistent form
        "KLAGENFURT AM WÖRTHERSEE":"KLAGENFURT AM WORTHERSEE",
        "KLAGENFURT":"KLAGENFURT AM WORTHERSEE",
        "ST. POLTEN":"ST. POLTEN","ST PÖLTEN":"ST. POLTEN","SANKT POLTEN":"ST. POLTEN",
        "WIENER NEUSTADT":"WIENER NEUSTADT",
        "WIENER NEUDORF":"WIENER NEUDORF",
        "TULLN AN DER DONAU":"TULLN AN DER DONAU","TULLN":"TULLN AN DER DONAU",
        "KREMS AN DER DONAU":"KREMS AN DER DONAU","KREMS":"KREMS AN DER DONAU",
        "VOSENDORF":"VOSENDORF","VÖSENDORF":"VOSENDORF",
        "MODLING":"MODLING","MÖDLING":"MODLING",
    },
    "CZ": {
        # Prague districts
        "PRAHA 1":"PRAHA","PRAHA 2":"PRAHA","PRAHA 3":"PRAHA",
        "PRAHA 4":"PRAHA","PRAHA 5":"PRAHA","PRAHA 6":"PRAHA",
        "PRAHA 7":"PRAHA","PRAHA 8":"PRAHA","PRAHA 9":"PRAHA",
        "PRAGUE 1":"PRAHA","PRAGUE 2":"PRAHA","PRAGUE 3":"PRAHA",
        "VINOHRADY":"PRAHA","ZIZKOV":"PRAHA","ŽIŽKOV":"PRAHA",
        "SMICHOV":"PRAHA","SMÍCHOV":"PRAHA","DEJVICE":"PRAHA",
        "HOLESOVICE":"PRAHA","HOLEŠOVICE":"PRAHA",
        "KARLIN":"PRAHA","KARLÍN":"PRAHA",
        "BRNO-STRED":"BRNO","BRNO-STŘED":"BRNO","BRNO MESTO":"BRNO","BRNO MĚSTO":"BRNO",
    },
    "HU": {
        # Budapest districts
        "PEST":"BUDAPEST","BUDA":"BUDAPEST","OBUDA":"BUDAPEST","ÓBUDA":"BUDAPEST",
        "UJBUDA":"BUDAPEST","ÚJBUDA":"BUDAPEST",
        "FERENCVAROS":"BUDAPEST","FERENCVÁROS":"BUDAPEST",
        "KOBANYA":"BUDAPEST","KŐBÁNYA":"BUDAPEST",
        "ZUGLO":"BUDAPEST","ZUGLÓ":"BUDAPEST",
        "UJPEST":"BUDAPEST","ÚJPEST":"BUDAPEST",
        "ANGYALFOLD":"BUDAPEST","ANGYALFÖLD":"BUDAPEST",
        "I. KERULET":"BUDAPEST","II. KERULET":"BUDAPEST","III. KERULET":"BUDAPEST",
        "IV. KERULET":"BUDAPEST","V. KERULET":"BUDAPEST","VI. KERULET":"BUDAPEST",
        "VII. KERULET":"BUDAPEST","VIII. KERULET":"BUDAPEST","IX. KERULET":"BUDAPEST",
        "X. KERULET":"BUDAPEST","XI. KERULET":"BUDAPEST","XII. KERULET":"BUDAPEST",
        "XIII. KERULET":"BUDAPEST","XIV. KERULET":"BUDAPEST","XV. KERULET":"BUDAPEST",
        "XVI. KERULET":"BUDAPEST","XVII. KERULET":"BUDAPEST","XVIII. KERULET":"BUDAPEST",
        "XIX. KERULET":"BUDAPEST","XX. KERULET":"BUDAPEST","XXI. KERULET":"BUDAPEST",
        "XXII. KERULET":"BUDAPEST","XXIII. KERULET":"BUDAPEST",
    },
}


def norm_city(s, char_map, market_code=""):
    """Normalise city name and resolve district → province for the given market."""
    if pd.isna(s) or not str(s).strip():
        return ""
    s = str(s).strip()
    for k, v in char_map.items():
        s = s.replace(k, v)
    s = re.sub(r"\s+(OD|MERKEZ|DISTRICT|ILCE)$", "", s, flags=re.IGNORECASE)
    s = s.upper().strip()
    district_map = _DISTRICT_MAP.get(market_code.upper(), {})
    return district_map.get(s, s)


def cities_compatible(city_a, city_b):
    """Return True if cities are the same, or either is unknown/blank (can't rule out)."""
    a = str(city_a).strip().upper()
    b = str(city_b).strip().upper()
    if not a or not b or a in ("UNKNOWN", "NAN") or b in ("UNKNOWN", "NAN"):
        return True   # can't verify — don't block
    return a == b


def load_zones(file=None, market_code=None):
    """
    Load delivery zone polygons. Two modes:
    1. Built-in: pass market_code ('NO' or 'TR') to load bundled zone file.
    2. Custom upload: pass a file object (CSV or XLSX) with WKT column.
    Returns a list of dicts: {polygon, zone_name, city_name, zone_id}
    WKT format: POLYGON ((lng lat, ...)) — longitude first, latitude second.
    """
    import os

    def _parse_df(df):
        zones = []
        wkt_col   = next((c for c in df.columns if "wkt" in str(c).lower()), None)
        zname_col = next((c for c in df.columns if "zone_name" in str(c).lower()), None)
        cname_col = next((c for c in df.columns if "city_name" in str(c).lower()), None)
        zid_col   = next((c for c in df.columns if "zone_id"   in str(c).lower()), None)
        if not wkt_col:
            return zones
        for _, row in df.iterrows():
            wkt_str = str(row.get(wkt_col, "")).strip()
            if not wkt_str or wkt_str.lower() == "nan":
                continue
            try:
                polygon = wkt_loads(wkt_str)
                zones.append({
                    "polygon":   polygon,
                    "zone_name": str(row.get(zname_col, "")) if zname_col else "",
                    "city_name": str(row.get(cname_col, "")) if cname_col else "",
                    "zone_id":   str(row.get(zid_col,   "")) if zid_col   else "",
                })
            except Exception:
                continue
        return zones

    # ── Custom upload takes priority ──────────────────────────
    if file is not None:
        try:
            df = pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)
            return _parse_df(df)
        except Exception:
            return []

    # ── Built-in bundled zones ─────────────────────────────────
    if market_code in ("NO", "TR", "SE"):
        # Look for the JSON file next to app.py
        base_dir  = os.path.dirname(os.path.abspath(__file__))
        json_path = os.path.join(base_dir, f"zones_{market_code}.json")
        if os.path.exists(json_path):
            try:
                with open(json_path, encoding="utf-8") as f:
                    raw = json.load(f)
                zones = []
                for z in raw:
                    try:
                        zones.append({
                            "polygon":   wkt_loads(z["wkt"]),
                            "zone_name": z.get("zone_name", ""),
                            "city_name": z.get("city_name", ""),
                            "zone_id":   z.get("zone_id",   ""),
                        })
                    except Exception:
                        continue
                return zones
            except Exception:
                return []

    return []


def geocode_address(street, city, postal_code, country_suffix, cache={}):
    """
    Geocode a street address using Nominatim (free, no API key).
    Returns (lat, lng) or (None, None).
    Caches results to avoid re-fetching identical addresses.
    Respects Nominatim's 1 req/sec rate limit.

    Handles Salesforce exports where Street may contain the full address
    (e.g. "Storgatan 5, 123 45 Stockholm, Sweden") or the "EMEA" country
    placeholder instead of the actual country name.
    """
    key = f"{street}|{city}|{postal_code}|{country_suffix}"
    if key in cache:
        return cache[key]

    import re as _re

    # ── Clean the street field ──────────────────────────────
    street_clean = str(street).strip() if street and str(street).strip() not in ("", "nan") else ""

    # Remove Salesforce "EMEA" country placeholder
    street_clean = _re.sub(r',?\s*EMEA\b', '', street_clean).strip().rstrip(',').strip()

    if not street_clean:
        cache[key] = (None, None)
        return (None, None)

    # Detect if street is already a full address:
    # has a postal code (4-6 digits) AND a recognisable country/city name
    _known = ["sweden", "norway", "turkey", "austria", "czech", "hungary",
              "sverige", "norge", "türkiye"]
    has_postal   = bool(_re.search(r'\b\d{4,6}\b', street_clean))
    has_country  = any(k in street_clean.lower() for k in _known)
    if country_suffix:
        has_country = has_country or country_suffix.lower() in street_clean.lower()

    if has_postal and has_country:
        # Street is self-contained — use it directly
        query = street_clean
    else:
        # Build from parts, avoid duplicating city already in the street
        city_clean = str(city).strip() if city and str(city).strip() not in ("", "nan") else ""
        zip_clean  = str(postal_code).strip() if postal_code and str(postal_code).strip() not in ("", "nan") else ""
        if city_clean and city_clean.lower() in street_clean.lower():
            city_clean = ""
        parts = [p for p in [street_clean, zip_clean, city_clean, country_suffix]
                 if p and p.strip() and p.strip() != "nan"]
        if not parts:
            cache[key] = (None, None)
            return (None, None)
        query = ", ".join(p.strip() for p in parts)

    params = urlencode({"q": query, "format": "json", "limit": "1"})
    url    = f"https://nominatim.openstreetmap.org/search?{params}"
    headers = {"User-Agent": "LeadClassifier/1.0"}

    try:
        time.sleep(1.1)  # Nominatim rate limit: 1 req/sec
        from urllib.request import Request
        req  = Request(url, headers=headers)
        resp = urlopen(req, timeout=6)
        data = json.loads(resp.read().decode())
        if data:
            lat = float(data[0]["lat"])
            lng = float(data[0]["lon"])
            cache[key] = (lat, lng)
            return (lat, lng)
    except Exception:
        pass

    cache[key] = (None, None)
    return (None, None)


def point_in_zones(lat, lng, zones):
    """
    Check if (lat, lng) falls inside any zone polygon.
    WKT uses (lng lat) order so we create Point(lng, lat).
    Returns (zone_name, city_name) of the first matching zone, or (None, None).
    """
    if lat is None or lng is None:
        return None, None
    try:
        pt = Point(float(lng), float(lat))  # WKT is (lng lat)
        for z in zones:
            if z["polygon"].contains(pt):
                return z["zone_name"], z["city_name"]
    except Exception:
        pass
    return None, None


def check_delivery_zone(row, col_map_leads, zones, country_suffix, geocode_enabled=True):
    """
    For a single lead row, determine delivery zone status.
    Returns (status, zone_name, city_name, method_used).
    Status: 'Within Zone' | 'Outside Zone' | 'No Zone Data' | 'Geocoding Failed'
    """
    if not zones:
        return "No Zone Data", "", "", ""

    lat_col  = col_map_leads.get("lat")
    lng_col  = col_map_leads.get("lng")
    str_col  = col_map_leads.get("street")
    city_col = col_map_leads.get("city")
    zip_col  = col_map_leads.get("zip")

    # Try coordinates first
    lat = row.get(lat_col) if lat_col else None
    lng = row.get(lng_col) if lng_col else None
    if lat is not None and lng is not None:
        try:
            lat, lng = float(lat), float(lng)
            if not (pd.isna(lat) or pd.isna(lng)):
                zone_name, city_name = point_in_zones(lat, lng, zones)
                if zone_name is not None:
                    return "Within Zone", zone_name, city_name, "Coordinates"
                else:
                    return "Outside Zone", "", "", "Coordinates"
        except (ValueError, TypeError):
            pass

    # Fall back to geocoding
    if geocode_enabled:
        street  = str(row.get(str_col,  "") or "") if str_col  else ""
        city    = str(row.get(city_col,  "") or "") if city_col else ""
        postal  = str(row.get(zip_col,   "") or "") if zip_col  else ""
        if any([street.strip(), city.strip(), postal.strip()]):
            lat, lng = geocode_address(street, city, postal, country_suffix)
            if lat is not None:
                zone_name, city_name = point_in_zones(lat, lng, zones)
                if zone_name is not None:
                    return "Within Zone", zone_name, city_name, "Geocoded"
                else:
                    return "Outside Zone", "", "", "Geocoded"
            return "Geocoding Failed", "", "", "Geocoded"

    return "Outside Zone", "", "", "No coordinates"


def classify_leads(leads_df, col_map_leads, crm_df, col_map_crm,
                   apify_df, col_map_apify, market_cfg,
                   confidence_threshold=0.5, zones=None, geocode_enabled=True):
    """Main classification pipeline."""
    char_map = market_cfg["char_map"]
    prefix   = market_cfg["phone_prefix"]

    # ── Build CRM lookups ──────────────────────────────────────────
    crm_phone_dict     = {}
    crm_name_city_dict = {}   # key: (name_norm, city_norm) — most precise
    crm_name_dict      = {}   # key: name_norm — fallback when CRM city is blank

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
            nn   = r.get("_name_norm", "")
            city = r.get("_city_norm", "")
            if nn:
                if city:
                    key = (nn, city)
                    if key not in crm_name_city_dict:
                        crm_name_city_dict[key] = r
                if nn not in crm_name_dict:
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
        lead_name   = row.get(name_col_l, "") if name_col_l else ""
        lead_street = row.get(street_col_l, "") if street_col_l else ""
        lead_city   = norm_city(row.get(city_col_l, ""), char_map, market_cfg.get("code", ""))

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
        crm_match    = None
        match_method = ""

        # 1. Phone match + name cross-check + city check
        if phone:
            candidate = crm_phone_dict.get(phone)
            if candidate is not None:
                crm_name_for_check = candidate.get(col_map_crm.get("name", ""), "")
                phone_name_score   = name_confidence(lead_name, crm_name_for_check, char_map)
                crm_cand_city      = candidate.get("_city_norm", "")
                city_ok            = cities_compatible(lead_city, crm_cand_city)
                if phone_name_score >= 0.45 and city_ok:
                    crm_match    = candidate
                    match_method = f"Phone + Name {phone_name_score:.2f}"
                # If city mismatch with good name score — silently skip (recycled number or wrong city)

        # 2. Exact name + city match (most precise)
        if crm_match is None:
            nn = norm_name(lead_name, char_map)
            if nn and lead_city:
                candidate = crm_name_city_dict.get((nn, lead_city))
                if candidate is not None:
                    crm_match    = candidate
                    match_method = "Name + City (exact)"

        # 3. Exact name only — verify city is compatible
        if crm_match is None:
            nn = norm_name(lead_name, char_map)
            if nn:
                candidate = crm_name_dict.get(nn)
                if candidate is not None:
                    crm_cand_city = candidate.get("_city_norm", "")
                    city_ok       = cities_compatible(lead_city, crm_cand_city)
                    if city_ok:
                        crm_match    = candidate
                        match_method = "Name (exact)" if crm_cand_city else "Name (exact, city unverified)"

        label = dup_grid = dup_name = dup_crm_city = dup_status = dup_reason = dup_method = ""
        if crm_match is not None:
            label        = "Duplicate"
            dup_grid     = str(crm_match.get(col_map_crm.get("grid", ""), ""))
            dup_name     = str(crm_match.get(col_map_crm.get("name", ""), ""))
            dup_crm_city = str(crm_match.get(col_map_crm.get("city", ""), "") or "")
            dup_status   = str(crm_match.get(col_map_crm.get("status", ""), ""))
            dup_reason   = str(crm_match.get(col_map_crm.get("reason", ""), "") or "")
            dup_method   = match_method

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
            match_conf = round(
                (0.4 if phone_hit else 0.0) +
                (min(name_score, 1.0) * 0.4) +
                (0.2 if addr_hit else 0.0), 3
            )
            # Phone match alone is always confirmed regardless of confidence score
            confirmed = phone_hit or (match_conf >= confidence_threshold)

            reasons = []
            if phone_hit: reasons.append("Phone ✓")
            reasons.append(f"Name {name_score:.2f}" + (" ✓" if name_score >= 0.5 else " ✗"))
            if addr_hit:  reasons.append("Address ✓")
            reasons.append(f"Confidence {match_conf:.2f}" + (" ✓" if confirmed else " ✗"))
            match_reason = " | ".join(reasons)

            if not label:
                if not confirmed:                        label = "Invalid Data"
                elif not gm_cat:                         label = "Invalid Data"
                elif perm or temp:                       label = "Business Closed"
                elif not is_food_delivery_eligible(gm_cat): label = "Wrong Target Group"
                else:                                    label = "Qualified / Convert"
        else:
            gm_biz_status = "Not Found on Google"
            match_reason  = "No Apify result"
            if not label:
                label = "Invalid Data"

        # ── Delivery zone check ────────────────────────────────
        zone_status = zone_name = zone_city = zone_method = ""
        if zones:
            zone_status, zone_name, zone_city, zone_method = check_delivery_zone(
                row, col_map_leads, zones,
                market_cfg.get("country_suffix", ""),
                geocode_enabled=geocode_enabled,
            )

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
            "Duplicate CRM City":    dup_crm_city,
            "CRM Account Status":    dup_status,
            "CRM Status Reason":     dup_reason,
            "Duplicate Match Method":dup_method,
            "Delivery Zone Status":  zone_status,
            "Zone Name":             zone_name,
            "Zone City":             zone_city,
            "Zone Method":           zone_method,
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

    # Col positions (1-based): GRID=1, Lead ID=2, Company=3, City=4, Street=5, Phone=6,
    # GM Title=7, GM Cat=8, GM BizStatus=9, GM Phone=10, GM Website=11, GM URL=12,
    # Match Conf=13, Match Reason=14, Label=15,
    # Dup GRID=16, Dup CRM Name=17, Dup CRM City=18, CRM Status=19, CRM Reason=20, Method=21
    LBL_R  = f"'Classified Leads'!O{DATA_S}:O{DATA_E}"
    CAT_R  = f"'Classified Leads'!H{DATA_S}:H{DATA_E}"
    CITY_R = f"'Classified Leads'!D{DATA_S}:D{DATA_E}"
    CRMS_R = f"'Classified Leads'!S{DATA_S}:S{DATA_E}"
    METH_R = f"'Classified Leads'!U{DATA_S}:U{DATA_E}"

    col_headers = [
        "GRID", "Lead ID", "Company / Account", "City", "Street", "Phone",
        "GM Title", "GM Category", "GM Business Status",
        "GM Phone", "GM Website", "GM URL",
        "Match Confidence", "Match Reason",
        "Label",
        "Duplicate GRID", "Duplicate CRM Name", "Duplicate CRM City",
        "CRM Account Status", "CRM Status Reason",
        "Duplicate Match Method",
        "Delivery Zone Status", "Zone Name", "Zone City", "Zone Method",
    ]

    # Check whether zone data is present in this run
    has_zones = df["Delivery Zone Status"].notna().any() and (df["Delivery Zone Status"] != "").any()

    # ── Sheet 1: All leads ───────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Classified Leads"
    ws1["A1"] = f"Lead Classification Report  |  {market_name}"
    ws1["A1"].font = Font(name="Arial", bold=True, size=14, color="1F4E79")
    ws1.merge_cells("A1:Y1")
    ws1["A2"] = "Total: {:,}   |   {}".format(
        len(df),
        "   ".join(f"{l}: {counts.get(l, 0):,}" for l in labels_order),
    )
    ws1["A2"].font = Font(name="Arial", italic=True, size=9, color="595959")
    ws1.merge_cells("A2:Y2")

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
            elif key == "Delivery Zone Status":
                zone_color  = "276221" if val == "Within Zone" else ("9C0006" if val == "Outside Zone" else "595959")
                zone_fill   = PatternFill("solid", start_color="C6EFCE") if val == "Within Zone" else \
                              (PatternFill("solid", start_color="FFC7CE") if val == "Outside Zone" else \
                               PatternFill("solid", start_color="EFEFEF"))
                c.font = Font(name="Arial", size=10, bold=True, color=zone_color)
                c.fill = zone_fill
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.font = Font(name="Arial", size=10)
                c.fill = fill
                c.alignment = Alignment(vertical="center")

    col_widths = [10, 18, 32, 14, 36, 16, 30, 24, 18, 16, 30, 50, 14, 38, 22, 14, 30, 14, 16, 18, 18, 18, 18, 14, 14]
    for i, w in enumerate(col_widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = "A5"
    ws1.auto_filter.ref = f"A4:Y{DATA_E}"

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

    # ── Zone summary (only if zone data present) ─────────────────
    if has_zones:
        ZONE_R = f"'Classified Leads'!V{DATA_S}:V{DATA_E}"
        ZN_R   = f"'Classified Leads'!W{DATA_S}:W{DATA_E}"
        sec(ws2, 47, 1, "DELIVERY ZONE BREAKDOWN", span=3)
        for ci, h in enumerate(["Zone Status", "Count", "% of Total"], 1):
            hdr(ws2, 48, ci, h)
        zone_statuses = [("Within Zone", "276221", "C6EFCE"),
                         ("Outside Zone", "9C0006", "FFC7CE"),
                         ("Geocoding Failed", "595959", "D9D9D9"),
                         ("No Zone Data", "595959", "EFEFEF")]
        for i, (zs, col, bg) in enumerate(zone_statuses):
            r    = 49 + i
            fill = PatternFill("solid", start_color=bg)
            dc(ws2, r, 1, zs, fill=fill, align="left", bold=True, color=col)
            dc(ws2, r, 2, f'=COUNTIF({ZONE_R},"{zs}")', fill=fill)
            dc(ws2, r, 3, f'=IF(B12=0,0,B{r}/B12)', fill=fill, fmt="0.0%")

        # Top zones
        sec(ws2, 47, 5, "TOP DELIVERY ZONES", span=3)
        for ci, h in enumerate(["Zone Name", "Within", "Outside"], 5):
            hdr(ws2, 48, ci, h)
        top_zones = [z for z in df["Zone Name"].value_counts().head(15).index.tolist() if z and z != ""]
        for i, zn in enumerate(top_zones):
            r    = 49 + i
            fill = PatternFill("solid", start_color="EBF3FB") if i % 2 == 0 else PatternFill("solid", start_color="FFFFFF")
            dc(ws2, r, 5, zn, fill=fill, bold=True, align="left")
            dc(ws2, r, 6, f'=COUNTIFS({ZONE_R},"Within Zone",{ZN_R},E{r})', fill=fill)
            dc(ws2, r, 7, f'=COUNTIFS({ZONE_R},"Outside Zone",{ZN_R},E{r})', fill=fill)

    for col, w in zip("ABCDE",  [28, 10, 12, 2, 22]):
        ws2.column_dimensions[col].width = w
    for col, w in zip("FGHIJ",  [14, 12, 12, 12, 10]):
        ws2.column_dimensions[col].width = w

    # ── Sheet 3: Qualified ───────────────────────────────────────
    ws3 = wb.create_sheet("✅ Qualified")
    q_df = df[df["Label"] == "Qualified / Convert"].reset_index(drop=True)
    ws3["A1"] = f"Qualified / Convert – Sales Ready  ({len(q_df):,} leads)"
    ws3["A1"].font = Font(name="Arial", bold=True, size=13, color="276221")
    ws3.merge_cells("A1:Q1")
    q_heads = ["GRID", "Lead ID", "Company / Account", "City", "Street", "Phone",
               "GM Title", "GM Category", "GM Phone", "GM Website", "GM URL",
               "Match Confidence", "Match Reason",
               "Delivery Zone Status", "Zone Name", "Zone City", "Zone Method"]
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
            elif key == "Delivery Zone Status":
                zone_color = "276221" if val == "Within Zone" else ("9C0006" if val == "Outside Zone" else "595959")
                zone_fill  = PatternFill("solid", start_color="C6EFCE") if val == "Within Zone" else \
                             (PatternFill("solid", start_color="FFC7CE") if val == "Outside Zone" else \
                              PatternFill("solid", start_color="EFEFEF"))
                c.font = Font(name="Arial", size=10, bold=True, color=zone_color)
                c.fill = zone_fill; c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.font = Font(name="Arial", size=10)
                c.fill = fill; c.alignment = Alignment(vertical="center")
    for i, w in enumerate([10, 18, 32, 14, 36, 16, 30, 24, 16, 30, 50, 14, 38, 18, 18, 14, 14], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.freeze_panes = "A4"
    ws3.auto_filter.ref = f"A3:Q{3 + len(q_df)}"

    # ── Sheet 4: Duplicates ──────────────────────────────────────
    ws4 = wb.create_sheet("🔴 Duplicates")
    d_df = df[df["Label"] == "Duplicate"].reset_index(drop=True)
    ws4["A1"] = f"Duplicates – Found in CRM  ({len(d_df):,} leads)"
    ws4["A1"].font = Font(name="Arial", bold=True, size=13, color="9C0006")
    ws4.merge_cells("A1:L1")
    d_heads = ["GRID", "Lead ID", "Company / Account", "City", "Phone",
               "Duplicate GRID", "Duplicate CRM Name", "Duplicate CRM City",
               "CRM Account Status", "CRM Status Reason",
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
    for i, w in enumerate([10, 18, 32, 14, 16, 14, 30, 14, 16, 18, 18, 24, 18], 1):
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
    if not check_password():
        return

    # ── Pandora brand CSS ──────────────────────────────────────
    st.markdown("""
    <style>
    /* ── Base ── */
    html, body, [class*="css"] {
        font-family: Arial, sans-serif;
    }

    /* ── Top header bar — subtle pink strip ── */
    header[data-testid="stHeader"] {
        background: #FFFFFF;
        border-bottom: 2px solid rgba(223,16,103,0.25);
    }

    /* ── Sidebar — clean white with pink left border ── */
    [data-testid="stSidebar"] {
        background: #FAFAFA;
        border-right: 2px solid rgba(223,16,103,0.3);
    }
    [data-testid="stSidebar"] * { color: #1A1A1A !important; }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #DF1067 !important;
        font-weight: 700 !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
        color: #444444 !important;
        font-size: 0.88rem !important;
        line-height: 1.55 !important;
    }
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stSlider label {
        color: #333333 !important;
        font-weight: 600 !important;
    }
    [data-testid="stSidebar"] .stSelectbox > div > div {
        background: #FFFFFF !important;
        border: 1.5px solid #E0E0E0 !important;
        border-radius: 6px !important;
        color: #1A1A1A !important;
    }
    [data-testid="stSidebar"] .stSelectbox > div > div:focus-within {
        border-color: #DF1067 !important;
    }
    [data-testid="stSidebar"] hr {
        border-color: #EBEBEB !important;
        margin: 0.8rem 0 !important;
    }

    /* ── Page title ── */
    .pandora-title {
        background: linear-gradient(90deg, #DF1067, #B8209D);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 2.2rem;
        font-weight: 700;
        margin-bottom: 0;
        line-height: 1.2;
    }
    .pandora-caption {
        color: #888888;
        font-size: 0.92rem;
        margin-top: 0.2rem;
        margin-bottom: 1.4rem;
    }

    /* ── Tabs ── */
    [data-testid="stTabs"] button {
        font-weight: 600;
        font-size: 0.92rem;
        color: #666666 !important;
        border-radius: 0 !important;
        padding-bottom: 0.6rem !important;
    }
    [data-testid="stTabs"] button:hover {
        color: #DF1067 !important;
        background: transparent !important;
    }
    [data-testid="stTabs"] button[aria-selected="true"] {
        color: #DF1067 !important;
        border-bottom: 3px solid #DF1067 !important;
        background: transparent !important;
    }

    /* ── Primary run button ── */
    [data-testid="stButton"] > button[kind="primary"] {
        background: #DF1067 !important;
        border: none !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.5rem !important;
        letter-spacing: 0.01em !important;
        transition: background 0.15s ease !important;
    }
    [data-testid="stButton"] > button[kind="primary"]:hover {
        background: #C00055 !important;
    }

    /* ── Download button ── */
    [data-testid="stDownloadButton"] > button {
        background: #DF1067 !important;
        border: none !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        width: 100% !important;
    }
    [data-testid="stDownloadButton"] > button:hover {
        background: #C00055 !important;
    }

    /* ── Link buttons — outlined style ── */
    [data-testid="stLinkButton"] > a {
        background: #FFFFFF !important;
        border: 1.5px solid #DF1067 !important;
        color: #DF1067 !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
        font-size: 0.88rem !important;
    }
    [data-testid="stLinkButton"] > a:hover {
        background: #FFF0F5 !important;
    }

    /* ── Metric cards — clean white with subtle border ── */
    [data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid #EBEBEB;
        border-top: 3px solid #DF1067;
        border-radius: 8px;
        padding: 0.9rem 1rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    [data-testid="stMetricLabel"] {
        color: #555555 !important;
        font-weight: 600 !important;
        font-size: 0.82rem !important;
        text-transform: uppercase !important;
        letter-spacing: 0.04em !important;
    }
    [data-testid="stMetricValue"] {
        color: #1A1A1A !important;
        font-weight: 700 !important;
        font-size: 1.8rem !important;
    }

    /* ── File uploader — clean minimal style ── */
    [data-testid="stFileUploader"] {
        border: 2px dashed #D0D0D0 !important;
        border-radius: 10px !important;
        background: #FAFAFA !important;
        transition: border-color 0.15s ease !important;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #DF1067 !important;
        background: #FFF8FB !important;
    }

    /* ── Expander / help section ── */
    [data-testid="stExpander"] summary {
        background: #FFFFFF !important;
        border: 1px solid #E8E8E8 !important;
        border-left: 3px solid #DF1067 !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        color: #333333 !important;
        font-size: 0.9rem !important;
    }
    [data-testid="stExpander"] summary:hover {
        background: #FFF8FB !important;
        border-left-color: #B8209D !important;
    }
    [data-testid="stExpander"] [data-testid="stExpanderDetails"] {
        border: 1px solid #E8E8E8 !important;
        border-top: none !important;
        border-radius: 0 0 6px 6px !important;
        background: #FFFFFF !important;
    }

    /* ── Alerts ── */
    [data-testid="stAlert"][data-type="success"] {
        background: #F6FFF9 !important;
        border-left: 4px solid #2E7D32 !important;
        color: #1B5E20 !important;
    }
    [data-testid="stAlert"][data-type="info"] {
        background: #F5F5FF !important;
        border-left: 4px solid #DF1067 !important;
    }
    [data-testid="stAlert"][data-type="warning"] {
        background: #FFFBF0 !important;
        border-left: 4px solid #F59E0B !important;
    }

    /* ── Dividers ── */
    hr { border-color: rgba(223,16,103,0.12) !important; }

    /* ── Toggle ── */
    [data-testid="stToggle"] span[data-checked="true"] {
        background: #DF1067 !important;
    }

    /* ── Spinner ── */
    [data-testid="stSpinner"] > div { border-top-color: #DF1067 !important; }

    /* ── Section headers in main area ── */
    .main h2 { color: #1A1A1A !important; font-size: 1.1rem !important; font-weight: 700 !important; }
    .main h3 { color: #444444 !important; font-size: 0.95rem !important; font-weight: 600 !important; }

    /* ── Caption / small text ── */
    [data-testid="stCaptionContainer"] { color: #888888 !important; font-size: 0.82rem !important; }

    /* ── Info box for zone status ── */
    .zone-badge-within  { display:inline-block; background:#E8F5E9; color:#2E7D32; border-radius:4px; padding:2px 8px; font-weight:600; font-size:0.85rem; }
    .zone-badge-outside { display:inline-block; background:#FFEBEE; color:#C62828; border-radius:4px; padding:2px 8px; font-weight:600; font-size:0.85rem; }

    /* ── Selectbox focus ── */
    [data-baseweb="select"] > div:focus-within { border-color: #DF1067 !important; }

    /* ── Slider ── */
    [data-testid="stSlider"] [data-testid="stThumbValue"] { color: #DF1067 !important; }
    </style>
    """, unsafe_allow_html=True)

    # ── Title ─────────────────────────────────────────────────
    st.markdown('<p class="pandora-title">🎯 Lead Classifier</p>', unsafe_allow_html=True)
    st.markdown('<p class="pandora-caption">Upload your lead file, Apify results and CRM export — get a classified Excel report.</p>', unsafe_allow_html=True)

    # ── Sidebar: Market selection ──────────────────────────────
    with st.sidebar:
        st.markdown(
            f'<div style="text-align:center; padding: 0.5rem 0 0.8rem 0;"><img src="data:image/png;base64,{DH_LOGO_B64}" style="width:120px;" /></div>',
            unsafe_allow_html=True
        )
        st.header("Settings")
        market_code = st.selectbox(
            "Market",
            options=[None] + list(MARKETS.keys()),
            format_func=lambda k: "— Select market —" if k is None else f"{MARKETS[k]['flag']} {MARKETS[k]['name']} ({k})",
        )
        market_cfg = MARKETS.get(market_code)
        if market_cfg:
            st.caption(f"Phone prefix: +{market_cfg['phone_prefix']}  ·  Country: {market_cfg['country_suffix']}")

        st.divider()
        st.subheader("Match confidence threshold")
        confidence_threshold = st.slider(
            "Minimum match confidence score",
            min_value=0.3,
            max_value=1.0,
            value=0.5,
            step=0.05,
            help="Leads where the overall Match Confidence score is below this value are marked Invalid Data. Confidence combines phone match (0.4), name similarity (0.4) and address match (0.2). Default 0.5 recommended.",
        )
        if confidence_threshold < 0.5:
            st.warning(f"⚠️ Threshold {confidence_threshold:.2f} — low confidence matches may include incorrect businesses.")
        elif confidence_threshold >= 0.8:
            st.info(f"Strict mode ({confidence_threshold:.2f}) — only strong phone + name matches will qualify.")
        else:
            st.success(f"Threshold: {confidence_threshold:.2f} (recommended)")

        st.divider()
        st.subheader("How it works")
        st.markdown("""
**Labels applied in this order:**

✅ **Qualified / Convert**
Not in CRM · confirmed on Google · food-delivery category · business open

🔴 **Duplicate**
Already exists in Salesforce — matched by phone, name and city

🟡 **Possible Duplicate**
Phone matched CRM but name looks different — check manually

🟡 **Business Closed**
Google shows permanently or temporarily closed

🟠 **Wrong Target Group**
Google category is not food-delivery eligible (hotel, hair salon, gym etc.)

⚫ **Invalid Data**
Not found on Google or confidence score too low
""")

        st.divider()
        st.subheader("Delivery zone check")
        st.markdown("""
**Built-in zones:** 🇳🇴 Norway · 🇹🇷 Turkey · 🇸🇪 Sweden

Zones load automatically when either market is selected. For other markets, upload a WKT zone file manually.

**How it works:**
- If the lead has **coordinates** → checked directly against zone polygons
- If **no coordinates** → address geocoded via OpenStreetMap (street + zip + city), then checked
- Result: ✅ Within Zone · 🚫 Outside Zone · ⚠️ Geocoding Failed

Zone results appear in the **Classified Leads** and **Qualified** tabs with green/red colour coding.

_Geocoding uses Nominatim (free, no API key) at 1 req/sec. Disable the toggle for faster runs when most leads have coordinates._
""")
        st.markdown("""
Three checks run in order — all require city to be compatible:

**1. Phone + Name + City**
Phone found in CRM → name similarity ≥ 0.45 → same city → Duplicate

**2. Name + City**
Exact name AND exact city both match CRM → Duplicate

**3. Name only**
Exact name matches CRM → city compatible or CRM has no city stored → Duplicate

**City check**
Districts are resolved to their city automatically — e.g. Södermalm → Stockholm, Kadıköy → Istanbul, Floridsdorf → Wien.

**Confidence score**
Phone match (0.4) + name similarity × 0.4 + location confirmation (0.2). Leads below the threshold you set are marked Invalid Data.
""")

    if not market_cfg:
        st.info("👈 Select a market from the sidebar to get started.")
        return

    # ── Tabs ──────────────────────────────────────────────────
    tab_classify, tab_urls = st.tabs(["📊 Classify leads", "🔗 Generate Apify URLs"])

    # ── TAB 1: Classify ───────────────────────────────────────
    with tab_classify:

        # ── Market-specific CRM report links ──────────────────
        CRM_REPORTS = {
            "NO": "https://deliveryhero.lightning.force.com/lightning/r/Report/00ObO0000047MEPUA2/view?queryScope=userFolders",
            "SE": "https://deliveryhero.lightning.force.com/lightning/r/Report/00ObO000004nwerUAA/view?queryScope=userFolders",
            "AT": "https://deliveryhero.lightning.force.com/lightning/r/Report/00ObO000004nxJBUAY/view?queryScope=userFolders",
            "HU": "https://deliveryhero.lightning.force.com/lightning/r/Report/00ObO000004nxO1UAI/view?queryScope=userFolders",
            "CZ": "https://deliveryhero.lightning.force.com/lightning/r/Report/00ObO000004nxW5UAI/view?queryScope=userFolders",
        }
        _crm_url = CRM_REPORTS.get(market_code)

        # ── Help links ────────────────────────────────────────
        with st.expander("📎 How to get your files — click to expand"):
            c1, c2, c3 = st.columns(3)

            with c1:
                st.markdown("**1. CRM export (Salesforce)**")
                if _crm_url:
                    st.markdown(
                        f"A dedicated report is set up for **{market_cfg['flag']} {market_cfg['name']}**. "
                        "Open it and export as CSV."
                    )
                    st.link_button(
                        f"Open {market_cfg['name']} CRM Report →",
                        _crm_url,
                        use_container_width=True,
                    )
                else:
                    st.markdown(
                        "Turkey has **more than 100k accounts** so the standard report is capped. "
                        "Use Salesforce Inspector with the SOQL query below to export the full CRM."
                    )
                    st.code(
                        "SELECT GRID__c, Name, Phone,\n"
                        "Account_Status__c, Status_Reason__c,\n"
                        "BillingCity\n"
                        "FROM Account\n"
                        "WHERE BillingCountry = 'Turkey'",
                        language="sql",
                    )
                    st.caption("Install the Chrome extension → open on any Salesforce page → Export tab → paste query → Download CSV.")
                    st.link_button(
                        "Salesforce Inspector Chrome Extension →",
                        "https://chromewebstore.google.com/detail/salesforce-inspector-reloaded/hpijlohoihegkfehhibggnkbjhoemldh",
                        use_container_width=True,
                    )

            with c2:
                st.markdown("**2. Leads export (Salesforce)**")
                st.markdown(
                    "Open the leads report, change the **Country** and **Lead Source** filters "
                    "to match your market, then export as CSV."
                )
                st.link_button(
                    "Open Leads Report →",
                    "https://deliveryhero.lightning.force.com/lightning/r/Report/00ObO0000047LqDUAU/view?queryScope=userFolders",
                    use_container_width=True,
                )

            with c3:
                st.markdown("**3. Apify — Google Maps Extractor**")
                st.markdown(
                    "Use the **URL generator tab** to create your search URLs first. "
                    "Paste them into Apify and run. Make sure these fields are selected:"
                )
                st.code(
                    "title, temporarilyClosed, permanentlyClosed,\n"
                    "postalCode, address, city, street,\n"
                    "website, phone, phoneUnformatted,\n"
                    "categoryName, categories, url, searchPageUrl",
                    language="text",
                )
                st.caption("Format: CSV · View: Overview · additionalInfo can be omitted.")
                st.link_button(
                    "Open Apify Actor →",
                    "https://console.apify.com/organization/ofPZhSPCC0KUPtU2Z/actors/nwua9Gu5YrADL7ZDj/input",
                    use_container_width=True,
                )

        st.divider()
        col1, col2, col3 = st.columns(3)

        with col1:
            st.subheader("1. Leads")
            leads_file = st.file_uploader("Upload leads file (.xlsx or .csv)", type=["xlsx", "csv"], key="leads")

        with col2:
            st.subheader("2. Apify Results")
            apify_file = st.file_uploader("Upload Apify Google Maps output (.csv or .xlsx)", type=["csv", "xlsx"], key="apify")
            st.caption("Optional — skipped if not uploaded.")

        with col3:
            st.subheader("3. CRM Export")
            crm_file = st.file_uploader("Upload Salesforce CRM export (.csv or .xlsx)", type=["csv", "xlsx"], key="crm")
            st.caption("Optional — duplicate check skipped if not uploaded.")

        st.divider()
        # Check if built-in zones exist for selected market
        import os as _os
        _base = _os.path.dirname(_os.path.abspath(__file__))
        _builtin_path = _os.path.join(_base, f"zones_{market_code}.json")
        _has_builtin  = _os.path.exists(_builtin_path)

        if _has_builtin:
            st.markdown(f"**📍 Delivery zones:** Built-in zones for **{market_cfg['flag']} {market_cfg['name']}** loaded automatically.")
            zone_file = st.file_uploader(
                "Upload custom zone file to override built-in zones (.csv or .xlsx) — optional",
                type=["csv", "xlsx"],
                key="zones",
                help="Leave empty to use the built-in delivery zones for this market."
            )
        else:
            st.markdown("**📍 Delivery zones:** No built-in zones for this market.")
            zone_file = st.file_uploader(
                "Upload delivery zone file (.csv or .xlsx) — optional",
                type=["csv", "xlsx"],
                key="zones",
                help="WKT polygon file from the logistics system. Each row is one delivery zone."
            )
            if not zone_file:
                st.caption("No zone file uploaded — delivery area check will be skipped.")

        geocode_toggle = st.toggle(
            "Geocode leads without coordinates",
            value=True,
            help="Uses OpenStreetMap Nominatim to find lat/lng from street + city + zip when coordinates are missing. Adds ~1 sec per lead. Disable for faster runs."
        )

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
            zones = []
            if zone_file:
                # Custom upload overrides built-in
                with st.spinner("Loading custom delivery zones..."):
                    zones = load_zones(file=zone_file)
                if zones:
                    st.info(f"📍 {len(zones)} zones loaded from uploaded file.")
                else:
                    st.warning("Zone file uploaded but no valid polygons found — check the WKT column.")
            elif _has_builtin:
                # Use built-in zones for this market
                with st.spinner(f"Loading built-in zones for {market_cfg['name']}..."):
                    zones = load_zones(market_code=market_code)
                if zones:
                    st.info(f"📍 {len(zones)} built-in delivery zones loaded for {market_cfg['flag']} {market_cfg['name']}.")
                else:
                    st.warning("Built-in zone file not found — delivery check skipped.")

            spinner_msg = "Classifying leads..."
            if zones and geocode_toggle:
                no_coords = sum(
                    1 for _, r in leads_df.iterrows()
                    if (not col_map_leads.get("lat") or pd.isna(r.get(col_map_leads.get("lat", ""), None)))
                )
                if no_coords > 0:
                    spinner_msg = f"Classifying leads (geocoding up to {no_coords} addresses — may take {no_coords}–{no_coords*2}s)..."

            with st.spinner(spinner_msg):
                result_df = classify_leads(
                    leads_df, col_map_leads,
                    crm_df, col_map_crm,
                    apify_df, col_map_apify,
                    market_cfg,
                    confidence_threshold=confidence_threshold,
                    zones=zones,
                    geocode_enabled=geocode_toggle,
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

            # ── Zone metrics (if zone file was used) ───────────
            if zones:
                zone_counts = result_df["Delivery Zone Status"].value_counts()
                within  = zone_counts.get("Within Zone", 0)
                outside = zone_counts.get("Outside Zone", 0)
                failed  = zone_counts.get("Geocoding Failed", 0)
                zc1, zc2, zc3 = st.columns(3)
                zc1.metric("📍 Within Delivery Zone",  within)
                zc2.metric("🚫 Outside Delivery Zone", outside)
                if failed:
                    zc3.metric("⚠️ Geocoding Failed", failed,
                               help="Address could not be geocoded — coordinates and address may be missing or too vague.")

            if crm_df is not None:
                dup_df      = result_df[result_df["Label"] == "Duplicate"]
                dup_methods = dup_df["Duplicate Match Method"].value_counts()
                phone_n     = sum(v for k, v in dup_methods.items() if "Phone" in str(k))
                name_city_n = dup_methods.get("Name + City (exact)", 0)
                name_only_n = sum(v for k, v in dup_methods.items() if "Name" in str(k) and "Phone" not in str(k)) - name_city_n
                st.caption(
                    f"Duplicates: {phone_n} by phone  ·  "
                    f"{name_city_n} by name + city  ·  "
                    f"{name_only_n} by name only"
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

                # Copy all URLs box
                valid_urls = [u for u in urls if u]
                st.markdown(f"**All URLs ({len(valid_urls)}) — select all and copy into Apify**")
                st.caption("Click inside the box → Ctrl+A (Win) / Cmd+A (Mac) → Copy")
                st.text_area(
                    label="urls",
                    value="\n".join(valid_urls),
                    height=200,
                    label_visibility="collapsed",
                )

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
                    "1. Copy the URLs above or download the CSV\n"
                    "2. Open Apify → Google Maps Extractor\n"
                    "3. Paste the URLs as input\n"
                    "4. Run the actor and download the results CSV\n"
                    "5. Come back to the **Classify leads** tab and upload it"
                )
            except Exception as e:
                st.error(f"Error: {e}")


if __name__ == "__main__":
    main()
