# ==========================================================================
#  EDIBLE OILS CHARTERER'S PLANNER
#  Перспектива фрахтователя: стоимость перевозки тонны vs спот
#  Запуск: python app.py  →  http://localhost:8050
# ==========================================================================
import io, json, math, heapq, os
from dash import Dash, html, dcc, Input, Output, State, ctx, dash_table
from dash.exceptions import PreventUpdate
import dash_leaflet as dl
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors as rlc
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm

# ── PALETTE ───────────────────────────────────────────────────────────────
C = {
    "bg":      "#e9edf1", "surface": "#ffffff", "surf2": "#dfe7ee",
    "ink":     "#15202b", "ink_l":   "#b9c7d3",
    "blue":    "#38648a", "blue_l":  "#c2d8e6", "blue_xl": "#e2edf4",
    "muted":   "#566877", "red":     "#8b2635", "red_l":   "#f3e2e5",
    "amber":   "#7a5230", "amber_l": "#f3e9da",
    "green":   "#2c5739", "green_l": "#daece0",
}
F = "-apple-system,'Segoe UI',Roboto,Arial,sans-serif"

def card(ch, **kw):
    return html.Div(ch, style={
        "background": C["surface"], "border": f"1.5px solid {C['ink']}",
        "borderRadius": "4px", "padding": "11px 13px", "marginBottom": "9px", **kw})

def lbl(t):
    return html.Div(t, style={
        "fontSize": "10.5px", "color": C["muted"], "marginTop": "6px",
        "marginBottom": "3px", "fontWeight": "700",
        "letterSpacing": "0.05em", "textTransform": "uppercase"})

def sec(t, color=None):
    return html.Div(t, style={
        "fontSize": "10.5px", "fontWeight": "800", "letterSpacing": "0.07em",
        "textTransform": "uppercase", "color": color or C["blue"],
        "borderBottom": f"1.5px solid {C['ink']}",
        "paddingBottom": "5px", "marginBottom": "8px"})

def kpi(title, value, sub="", color=None):
    return html.Div([
        html.Div(title, style={"fontSize": "10px", "color": C["muted"],
            "fontWeight": "700", "letterSpacing": "0.05em",
            "textTransform": "uppercase"}),
        html.Div(value, style={"fontSize": "17px", "fontWeight": "800",
            "color": color or C["ink"], "marginTop": "2px",
            "fontVariantNumeric": "tabular-nums"}),
        html.Div(sub, style={"fontSize": "10.5px", "color": C["muted"],
            "marginTop": "1px"}),
    ], style={"flex": "1", "padding": "9px 11px", "background": C["surface"],
        "border": f"1.5px solid {C['ink']}", "borderRadius": "4px",
        "minWidth": "130px"})

inp  = {"width": "100%", "padding": "5px 7px", "boxSizing": "border-box",
        "border": f"1.5px solid {C['ink_l']}", "borderRadius": "3px",
        "fontSize": "12px", "fontFamily": F, "background": C["surface"]}
btn  = {"padding": "7px 11px", "cursor": "pointer",
        "background": C["blue"], "color": "#fff",
        "border": f"1.5px solid {C['ink']}", "borderRadius": "3px",
        "fontWeight": "700", "fontSize": "12px", "fontFamily": F,
        "width": "100%", "marginTop": "4px"}
btng = {**btn, "background": C["surface"], "color": C["blue"]}
money = lambda x: "$" + f"{float(x):,.0f}"

TBL = {
    "style_table": {"overflowX": "auto",
        "border": f"1.5px solid {C['ink']}", "borderRadius": "3px"},
    "style_header": {"backgroundColor": C["blue"], "color": "#fff",
        "fontWeight": "700", "fontSize": "11px",
        "border": f"1px solid {C['ink']}"},
    "style_cell": {"backgroundColor": C["surface"], "color": C["ink"],
        "fontSize": "11px", "padding": "5px 8px",
        "border": f"1px solid {C['ink_l']}", "fontFamily": F,
        "minWidth": "80px", "maxWidth": "200px",
        "overflow": "hidden", "textOverflow": "ellipsis"},
    "style_data_conditional": [
        {"if": {"row_index": "odd"}, "backgroundColor": C["surf2"]},
    ],
}

# ── GEO NODES ─────────────────────────────────────────────────────────────
#  [lat, lon, type]  j=junction  p=port
NODES = {
    # junction / routing waypoints
    "Bishop Rock":          [49.750,  -6.583, "j"],
    "Cape Leeuwin":         [-34.533, 115.133,"j"],
    "Cape of Good Hope":    [-34.367,  18.383,"j"],
    "Fastnet Rock":         [51.333,  -9.600, "j"],
    "Grand Banks South":    [42.500, -50.000, "j"],
    "Honshu":               [35.000, 140.500, "j"],
    "Ile d Ouessant":       [48.667,  -5.500, "j"],
    "Inishtrahull":         [55.417,  -7.500, "j"],
    "Montreal":             [45.500, -73.550, "j"],
    "Nord-Ostsee-Kanal":    [54.367,  10.150, "j"],
    "Panama":               [ 8.883, -79.517, "j"],
    "Pentland Firth":       [58.700,  -3.333, "j"],
    "Port Said":            [31.267,  32.317, "j"],
    "Punta Arenas":         [-53.167,-70.900, "j"],
    "Selat Lombok":         [-8.833, 115.717, "j"],
    "Selat Sunda":          [-6.067, 105.833, "j"],
    "Selat Wetar":          [-8.317, 127.450, "j"],
    "Singapore":            [ 1.267, 103.833, "j"],
    "Skagens Odde":         [57.800,  10.733, "j"],
    "Strait of Gibraltar":  [35.950,  -5.750, "j"],
    "Straits of Florida":   [24.417, -83.000, "j"],
    "Torres Strait":        [-10.550,142.133, "j"],
    "Tsugaru-Kaikyo":       [41.500, 140.667, "j"],
    "Wilson Promontory":    [-39.167,146.433, "j"],
    "Yucatan Channel":      [21.833, -85.050, "j"],

    # Europe — Baltic / North Sea
    "St. Petersburg":       [59.950,  30.317, "p"],
    "Gdansk":               [54.350,  18.650, "p"],
    "Hamburg":              [53.550,   9.967, "p"],
    "Bremen":               [53.083,   8.800, "p"],
    "Rotterdam":            [51.900,   4.467, "p"],
    "Amsterdam":            [52.370,   4.900, "p"],
    "Antwerp":              [51.217,   4.400, "p"],
    "Ghent":                [51.050,   3.717, "p"],
    "Dunkirk":              [51.050,   2.367, "p"],
    "Le Havre":             [49.483,   0.117, "p"],
    "London":               [51.450,   0.367, "p"],
    "Liverpool":            [53.400,  -3.000, "p"],
    "Bilbao":               [43.350,  -3.050, "p"],

    # Europe — Med / Black Sea
    "Algeciras":            [36.133,  -5.433, "p"],
    "Tanger Med":           [35.883,  -5.500, "p"],
    "Lisbon":               [38.700,  -9.133, "p"],
    "Casablanca":           [33.617,  -7.617, "p"],
    "Las Palmas":           [28.142, -15.417, "p"],
    "Valencia":             [39.450,  -0.333, "p"],
    "Barcelona":            [41.367,   2.183, "p"],
    "Marseille":            [43.300,   5.367, "p"],
    "Genoa":                [44.400,   8.933, "p"],
    "Livorno":              [43.550,  10.300, "p"],
    "Piraeus":              [37.950,  23.617, "p"],
    "Thessaloniki":         [40.633,  22.933, "p"],
    "Istanbul":             [41.017,  28.967, "p"],
    "Izmir":                [38.417,  27.133, "p"],
    "Mersin":               [36.800,  34.633, "p"],
    "Iskenderun":           [36.583,  36.167, "p"],
    "Constanta":            [44.167,  28.650, "p"],
    "Burgas":               [42.500,  27.467, "p"],
    "Odessa":               [46.483,  30.733, "p"],
    "Yuzhny":               [46.617,  31.000, "p"],
    "Novorossiysk":         [44.733,  37.783, "p"],
    "Taman":                [45.200,  36.700, "p"],   # ← NEW

    # Russia Arctic
    "Vitino":               [67.283,  32.350, "p"],   # ← NEW (Murmansk Oblast, White Sea)

    # Levant / Red Sea
    "Alexandria":           [31.200,  29.883, "p"],
    "Damietta":             [31.417,  31.817, "p"],
    "Beirut":               [33.900,  35.517, "p"],
    "Haifa":                [32.833,  34.983, "p"],
    "Jeddah":               [21.483,  39.167, "p"],
    "Yanbu":                [24.083,  38.033, "p"],
    "Djibouti":             [11.600,  43.150, "p"],
    "Aden":                 [12.794,  44.983, "p"],

    # Gulf / Pakistan
    "Fujairah":             [25.117,  56.333, "p"],
    "Khor Fakkan":          [25.333,  56.350, "p"],
    "Jebel Ali":            [25.007,  55.063, "p"],
    "Sharjah":              [25.333,  55.400, "p"],
    "Bandar Abbas":         [27.183,  56.283, "p"],
    "Karachi":              [24.800,  66.983, "p"],
    "Port Qasim":           [24.783,  67.350, "p"],

    # India West Coast
    "Mumbai":               [18.933,  72.850, "p"],
    "Kandla":               [23.033,  70.217, "p"],
    "Mundra":               [22.840,  69.717, "p"],
    "New Mangalore":        [12.950,  74.800, "p"],   # ← NEW

    # India East Coast
    "Chennai":              [13.083,  80.267, "p"],
    "Krishnapatnam":        [14.250,  80.130, "p"],   # ← NEW
    "Kakinada":             [16.933,  82.233, "p"],   # ← NEW
    "Tuticorin":            [ 8.800,  78.150, "p"],
    "Kolkata":              [22.567,  88.317, "p"],
    "Haldia":               [22.033,  88.083, "p"],
    "Chittagong":           [22.333,  91.833, "p"],
    "Colombo":              [ 6.950,  79.850, "p"],

    # SE Asia
    "Port Klang":           [ 3.000, 101.400, "p"],
    "Westport":             [ 3.067, 101.367, "p"],
    "Pasir Gudang":         [ 1.467, 103.900, "p"],
    "Tanjung Pelepas":      [ 1.360, 103.550, "p"],
    "Penang":               [ 5.400, 100.317, "p"],
    "Belawan":              [ 3.783,  98.683, "p"],
    "Dumai":                [ 1.667, 101.467, "p"],
    "Kuala Tanjung":        [ 2.967,  99.450, "p"],
    "Tanjung Priok":        [-6.100, 106.883, "p"],
    "Panjang":              [-5.467, 105.317, "p"],
    "Cigading":             [-6.017, 106.000, "p"],
    "Surabaya":             [-7.200, 112.733, "p"],
    "Bangkok":              [13.733, 100.500, "p"],
    "Laem Chabang":         [13.083, 100.900, "p"],
    "Ho Chi Minh":          [10.750, 106.750, "p"],
    "Manila":               [14.600, 120.950, "p"],

    # China / NE Asia
    "Hong Kong":            [22.300, 114.167, "p"],
    "Nansha":               [22.717, 113.600, "p"],
    "Guangzhou":            [23.083, 113.467, "p"],
    "Xiamen":               [24.450, 118.067, "p"],
    "Shanghai":             [31.233, 121.492, "p"],
    "Ningbo":               [29.867, 122.133, "p"],
    "Qingdao":              [36.067, 120.317, "p"],
    "Tianjin":              [38.983, 117.733, "p"],
    "Dalian":               [38.917, 121.633, "p"],
    "Kaohsiung":            [22.617, 120.283, "p"],
    "Busan":                [35.100, 129.050, "p"],
    "Ulsan":                [35.500, 129.383, "p"],
    "Gwangyang":            [34.917, 127.700, "p"],
    "Incheon":              [37.450, 126.633, "p"],
    "Kobe":                 [34.667, 135.200, "p"],
    "Nagoya":               [35.050, 136.883, "p"],
    "Yokohama":             [35.450, 139.650, "p"],
    "Tokyo":                [35.633, 139.783, "p"],
    "Vladivostok":          [43.108, 131.883, "p"],

    # Americas — East
    "New York":             [40.700, -74.017, "p"],
    "Philadelphia":         [40.000, -75.133, "p"],
    "Baltimore":            [39.283, -76.600, "p"],
    "Norfolk":              [36.850, -76.300, "p"],
    "Savannah":             [32.083, -81.100, "p"],
    "Miami":                [25.767, -80.183, "p"],
    "New Orleans":          [29.950, -90.067, "p"],
    "Houston":              [29.733, -95.017, "p"],
    "Corpus Christi":       [27.800, -97.400, "p"],
    "Colon":                [ 9.367, -79.917, "p"],
    "Veracruz":             [19.200, -96.133, "p"],

    # Americas — West
    "Los Angeles":          [33.750,-118.271, "p"],
    "Long Beach":           [33.767,-118.200, "p"],
    "Seattle":              [47.600,-122.333, "p"],
    "Vancouver":            [49.283,-123.117, "p"],
    "Valparaiso":           [-33.033,-71.633, "p"],
    "Callao":               [-12.050,-77.133, "p"],

    # South America — East
    "Santos":               [-23.950,-46.317, "p"],
    "Paranagua":            [-25.500,-48.517, "p"],
    "Rio Grande":           [-32.033,-52.100, "p"],
    "Montevideo":           [-34.900,-56.217, "p"],
    "Buenos Aires":         [-34.583,-58.367, "p"],
    "Rosario":              [-32.950,-60.650, "p"],
    "Bahia Blanca":         [-38.717,-62.267, "p"],

    # Africa
    "Dakar":                [14.683, -17.433, "p"],
    "Abidjan":              [ 5.300,  -4.017, "p"],
    "Tema":                 [ 5.633,   0.017, "p"],
    "Lagos":                [ 6.433,   3.400, "p"],
    "Douala":               [ 4.050,   9.700, "p"],
    "Cape Town":            [-33.900,  18.433,"p"],
    "Durban":               [-29.867,  31.033,"p"],
    "Richards Bay":         [-28.783,  32.033,"p"],
    "Mombasa":              [-4.067,   39.667,"p"],
    "Dar es Salaam":        [-6.817,   39.317,"p"],

    # Australia
    "Fremantle":            [-32.050, 115.739,"p"],
    "Adelaide":             [-34.850, 138.500,"p"],
    "Melbourne":            [-37.817, 144.967,"p"],
    "Sydney":               [-33.858, 151.217,"p"],
    "Brisbane":             [-27.367, 153.167,"p"],
    "Darwin":               [-12.467, 130.850,"p"],
    "Port Hedland":         [-20.317, 118.567,"p"],
    "Honolulu":             [21.308, -157.871,"p"],
}

# ── EDGES ─────────────────────────────────────────────────────────────────
EDGES = [
    # ── Pub.151 junction backbone
    ("Strait of Gibraltar","Bishop Rock",969),
    ("Strait of Gibraltar","Fastnet Rock",1068),
    ("Strait of Gibraltar","Grand Banks South",2070),
    ("Strait of Gibraltar","Ile d Ouessant",918),
    ("Strait of Gibraltar","Inishtrahull",1350),
    ("Strait of Gibraltar","Panama",4351),
    ("Strait of Gibraltar","Pentland Firth",1598),
    ("Strait of Gibraltar","Port Said",1943),
    ("Strait of Gibraltar","Cape of Good Hope",5082),
    ("Strait of Gibraltar","Punta Arenas",6352),
    ("Strait of Gibraltar","Straits of Florida",4009),
    ("Bishop Rock","Fastnet Rock",149),
    ("Bishop Rock","Grand Banks South",1830),
    ("Bishop Rock","Ile d Ouessant",80),
    ("Bishop Rock","Inishtrahull",395),
    ("Bishop Rock","Montreal",2844),
    ("Bishop Rock","Nord-Ostsee-Kanal",727),
    ("Bishop Rock","Panama",4388),
    ("Bishop Rock","Pentland Firth",636),
    ("Bishop Rock","Punta Arenas",7019),
    ("Bishop Rock","Skagens Odde",851),
    ("Bishop Rock","Straits of Florida",3859),
    ("Bishop Rock","Yucatan Channel",4176),
    ("Fastnet Rock","Cape of Good Hope",5880),
    ("Fastnet Rock","Grand Banks South",1714),
    ("Fastnet Rock","Ile d Ouessant",225),
    ("Fastnet Rock","Inishtrahull",335),
    ("Fastnet Rock","Panama",4247),
    ("Fastnet Rock","Punta Arenas",7051),
    ("Fastnet Rock","Straits of Florida",3761),
    ("Fastnet Rock","Yucatan Channel",4058),
    ("Grand Banks South","Cape of Good Hope",5954),
    ("Grand Banks South","Ile d Ouessant",1877),
    ("Grand Banks South","Montreal",1335),
    ("Grand Banks South","Panama",2555),
    ("Grand Banks South","Pentland Firth",1998),
    ("Grand Banks South","Straits of Florida",1990),
    ("Ile d Ouessant","Inishtrahull",329),
    ("Ile d Ouessant","Montreal",2901),
    ("Ile d Ouessant","Nord-Ostsee-Kanal",716),
    ("Ile d Ouessant","Panama",4374),
    ("Ile d Ouessant","Pentland Firth",872),
    ("Ile d Ouessant","Punta Arenas",6986),
    ("Ile d Ouessant","Skagens Odde",833),
    ("Inishtrahull","Montreal",2545),
    ("Inishtrahull","Pentland Firth",272),
    ("Inishtrahull","Punta Arenas",7315),
    ("Montreal","Panama",3204),
    ("Montreal","Pentland Firth",2641),
    ("Montreal","Straits of Florida",2531),
    ("Pentland Firth","Nord-Ostsee-Kanal",565),
    ("Pentland Firth","Skagens Odde",450),
    ("Skagens Odde","Nord-Ostsee-Kanal",224),
    ("Port Said","Cape of Good Hope",5340),
    ("Port Said","Cape Leeuwin",6389),
    ("Port Said","Selat Lombok",5892),
    ("Port Said","Selat Sunda",5224),
    ("Port Said","Singapore",5035),
    ("Port Said","Torres Strait",7423),
    ("Cape of Good Hope","Cape Leeuwin",4660),
    ("Cape of Good Hope","Selat Lombok",5461),
    ("Cape of Good Hope","Panama",6466),
    ("Cape of Good Hope","Punta Arenas",4262),
    ("Cape of Good Hope","Singapore",5579),
    ("Cape of Good Hope","Selat Sunda",5164),
    ("Cape of Good Hope","Torres Strait",6854),
    ("Cape of Good Hope","Selat Wetar",6768),
    ("Cape of Good Hope","Wilson Promontory",5530),
    ("Cape of Good Hope","Yucatan Channel",6770),
    ("Cape of Good Hope","Straits of Florida",6770),
    ("Cape Leeuwin","Selat Lombok",1571),
    ("Cape Leeuwin","Selat Sunda",1815),
    ("Cape Leeuwin","Selat Wetar",1954),
    ("Cape Leeuwin","Torres Strait",2717),
    ("Cape Leeuwin","Wilson Promontory",1524),
    ("Selat Lombok","Singapore",963),
    ("Selat Lombok","Selat Sunda",677),
    ("Selat Lombok","Selat Wetar",727),
    ("Selat Lombok","Honshu",3059),
    ("Selat Sunda","Singapore",532),
    ("Selat Sunda","Selat Wetar",1330),
    ("Selat Sunda","Honshu",3171),
    ("Selat Wetar","Singapore",1587),
    ("Selat Wetar","Torres Strait",881),
    ("Selat Wetar","Honshu",2838),
    ("Singapore","Honshu",2879),
    ("Singapore","Tsugaru-Kaikyo",3343),
    ("Singapore","Panama",10505),
    ("Honshu","Panama",7614),
    ("Honshu","Punta Arenas",9286),
    ("Honshu","Torres Strait",3265),
    ("Honshu","Tsugaru-Kaikyo",484),
    ("Honshu","Wilson Promontory",4881),
    ("Tsugaru-Kaikyo","Panama",8004),
    ("Tsugaru-Kaikyo","Punta Arenas",9449),
    ("Torres Strait","Panama",8451),
    ("Torres Strait","Punta Arenas",7217),
    ("Torres Strait","Wilson Promontory",2183),
    ("Wilson Promontory","Panama",7842),
    ("Wilson Promontory","Punta Arenas",5820),
    ("Panama","Yucatan Channel",855),
    ("Panama","Punta Arenas",3932),
    ("Straits of Florida","Yucatan Channel",192),

    # ── Europe ports
    ("Rotterdam","Bishop Rock",454),
    ("Rotterdam","Ile d Ouessant",444),
    ("Rotterdam","Nord-Ostsee-Kanal",323),
    ("Rotterdam","Pentland Firth",495),
    ("Rotterdam","Skagens Odde",447),
    ("Rotterdam","Strait of Gibraltar",1371),
    ("Amsterdam","Rotterdam",65),
    ("Antwerp","Rotterdam",55),
    ("Ghent","Antwerp",30),
    ("Le Havre","Bishop Rock",360),
    ("Le Havre","Ile d Ouessant",168),
    ("Dunkirk","Rotterdam",130),
    ("London","Bishop Rock",414),
    ("London","Rotterdam",200),
    ("Liverpool","Fastnet Rock",260),
    ("Bilbao","Bishop Rock",420),
    ("Hamburg","Nord-Ostsee-Kanal",127),
    ("Hamburg","Skagens Odde",305),
    ("Hamburg","Rotterdam",280),
    ("Bremen","Hamburg",130),
    ("Gdansk","Skagens Odde",530),
    ("Gdansk","Nord-Ostsee-Kanal",400),
    ("St. Petersburg","Skagens Odde",950),
    ("St. Petersburg","Nord-Ostsee-Kanal",800),

    # ── Vitino (White Sea oil terminal, Murmansk Oblast)
    ("Vitino","St. Petersburg",1380),     # White Sea → Gulf of Finland
    ("Vitino","Skagens Odde",2100),       # Arctic around Norway

    # ── Black Sea / Med
    ("Constanta","Istanbul",200),
    ("Odessa","Istanbul",340),
    ("Yuzhny","Odessa",20),
    ("Novorossiysk","Istanbul",500),
    ("Taman","Novorossiysk",50),          # Kerch Strait area
    ("Taman","Istanbul",430),
    ("Burgas","Istanbul",160),
    ("Istanbul","Piraeus",340),
    ("Izmir","Piraeus",190),
    ("Thessaloniki","Piraeus",300),
    ("Piraeus","Port Said",590),
    ("Piraeus","Strait of Gibraltar",1050),
    ("Barcelona","Strait of Gibraltar",625),
    ("Barcelona","Marseille",195),
    ("Valencia","Barcelona",180),
    ("Algeciras","Strait of Gibraltar",10),
    ("Tanger Med","Strait of Gibraltar",20),
    ("Genoa","Marseille",185),
    ("Livorno","Genoa",90),
    ("Lisbon","Strait of Gibraltar",320),
    ("Casablanca","Strait of Gibraltar",175),
    ("Casablanca","Las Palmas",515),
    ("Las Palmas","Strait of Gibraltar",681),
    ("Las Palmas","Bishop Rock",1360),
    ("Las Palmas","Straits of Florida",3713),

    # ── Levant / Red Sea
    ("Alexandria","Port Said",180),
    ("Damietta","Port Said",40),
    ("Beirut","Port Said",265),
    ("Haifa","Port Said",260),
    ("Mersin","Port Said",280),
    ("Iskenderun","Mersin",120),
    ("Jeddah","Port Said",660),
    ("Yanbu","Jeddah",200),
    ("Djibouti","Aden",200),
    ("Aden","Port Said",1400),
    ("Aden","Cape of Good Hope",3954),
    ("Aden","Singapore",3631),

    # ── Gulf / South Asia
    ("Fujairah","Aden",570),
    ("Khor Fakkan","Fujairah",15),
    ("Jebel Ali","Fujairah",70),
    ("Sharjah","Jebel Ali",30),
    ("Bandar Abbas","Fujairah",130),
    ("Karachi","Fujairah",660),
    ("Port Qasim","Karachi",25),
    ("Mumbai","Karachi",532),
    ("Kandla","Mumbai",390),
    ("Mundra","Kandla",60),

    # ── India West Coast
    ("New Mangalore","Mumbai",340),       # ← NEW
    ("New Mangalore","Colombo",450),      # ← NEW
    ("Tuticorin","Chennai",270),
    ("Tuticorin","Colombo",140),

    # ── India East Coast
    ("Chennai","Mumbai",790),
    ("Krishnapatnam","Chennai",170),      # ← NEW
    ("Krishnapatnam","Colombo",500),      # ← NEW
    ("Kakinada","Chennai",360),           # ← NEW
    ("Kakinada","Kolkata",540),           # ← NEW
    ("Kolkata","Chennai",870),
    ("Haldia","Kolkata",60),
    ("Chittagong","Kolkata",390),
    ("Colombo","Chennai",175),
    ("Colombo","Singapore",1581),
    ("Colombo","Port Said",3481),
    ("Colombo","Cape of Good Hope",4317),

    # ── SE Asia
    ("Port Klang","Singapore",200),
    ("Westport","Port Klang",15),
    ("Pasir Gudang","Singapore",30),
    ("Tanjung Pelepas","Singapore",40),
    ("Penang","Port Klang",230),
    ("Belawan","Singapore",240),
    ("Dumai","Singapore",150),
    ("Kuala Tanjung","Dumai",80),
    ("Tanjung Priok","Singapore",560),
    ("Panjang","Tanjung Priok",150),
    ("Cigading","Tanjung Priok",100),
    ("Surabaya","Tanjung Priok",420),
    ("Bangkok","Singapore",835),
    ("Laem Chabang","Bangkok",80),
    ("Ho Chi Minh","Singapore",650),
    ("Manila","Hong Kong",640),

    # ── China / NE Asia
    ("Hong Kong","Singapore",1460),
    ("Nansha","Hong Kong",110),
    ("Guangzhou","Hong Kong",130),
    ("Xiamen","Hong Kong",300),
    ("Shanghai","Hong Kong",770),
    ("Ningbo","Shanghai",130),
    ("Qingdao","Shanghai",390),
    ("Tianjin","Shanghai",790),
    ("Dalian","Tianjin",250),
    ("Kaohsiung","Hong Kong",340),
    ("Busan","Hong Kong",780),
    ("Busan","Honshu",450),
    ("Ulsan","Busan",50),
    ("Gwangyang","Busan",70),
    ("Incheon","Busan",450),
    ("Kobe","Busan",380),
    ("Kobe","Honshu",50),
    ("Nagoya","Kobe",130),
    ("Nagoya","Honshu",100),
    ("Yokohama","Honshu",20),
    ("Yokohama","Nagoya",200),
    ("Tokyo","Yokohama",15),
    ("Shanghai","Honshu",550),
    ("Vladivostok","Honshu",620),

    # ── Americas East
    ("New York","Strait of Gibraltar",3180),
    ("New York","Panama",2016),
    ("New York","Grand Banks South",950),
    ("New York","Straits of Florida",1087),
    ("Philadelphia","New York",80),
    ("Baltimore","New York",170),
    ("Norfolk","New York",280),
    ("Savannah","Norfolk",560),
    ("Miami","Savannah",680),
    ("New Orleans","Straits of Florida",900),
    ("Houston","New Orleans",300),
    ("Houston","Straits of Florida",1200),
    ("Houston","Panama",1150),
    ("Corpus Christi","Houston",200),
    ("Veracruz","Houston",750),
    ("Colon","Panama",44),

    # ── Americas West
    ("Los Angeles","Panama",2913),
    ("Long Beach","Los Angeles",15),
    ("Seattle","Vancouver",85),
    ("Vancouver","Los Angeles",979),
    ("Valparaiso","Punta Arenas",1500),
    ("Valparaiso","Panama",2983),
    ("Callao","Valparaiso",1274),
    ("Callao","Panama",1310),

    # ── South America East
    ("Santos","Strait of Gibraltar",4690),
    ("Santos","Cape of Good Hope",3560),
    ("Paranagua","Santos",180),
    ("Rio Grande","Santos",620),
    ("Buenos Aires","Santos",1080),
    ("Montevideo","Buenos Aires",120),
    ("Rosario","Buenos Aires",220),
    ("Bahia Blanca","Buenos Aires",420),
    ("Buenos Aires","Punta Arenas",1394),
    ("Buenos Aires","Cape of Good Hope",3704),

    # ── Africa
    ("Dakar","Las Palmas",930),
    ("Dakar","Strait of Gibraltar",1460),
    ("Abidjan","Dakar",950),
    ("Tema","Abidjan",220),
    ("Lagos","Tema",250),
    ("Douala","Lagos",350),
    ("Cape Town","Cape of Good Hope",35),
    ("Durban","Cape of Good Hope",764),
    ("Richards Bay","Durban",90),
    ("Mombasa","Durban",1390),
    ("Mombasa","Aden",1450),
    ("Dar es Salaam","Mombasa",220),

    # ── Australia
    ("Fremantle","Cape Leeuwin",181),
    ("Fremantle","Cape of Good Hope",4755),
    ("Fremantle","Selat Lombok",1440),
    ("Fremantle","Port Said",6310),
    ("Port Hedland","Fremantle",870),
    ("Darwin","Singapore",1840),
    ("Adelaide","Fremantle",1130),
    ("Melbourne","Adelaide",430),
    ("Sydney","Melbourne",560),
    ("Sydney","Torres Strait",1670),
    ("Sydney","Wilson Promontory",440),
    ("Brisbane","Sydney",440),
]

# ── ZONES ─────────────────────────────────────────────────────────────────
SEASONAL = [
    {"n":"North Atlantic Winter","m":[11,12,1,2,3],
     "a":40,"b":65,"c":-65,"d":20,"f":0.87},
    {"n":"North Sea Winter","m":[10,11,12,1,2,3],
     "a":48,"b":65,"c":-5,"d":15,"f":0.88},
    {"n":"Bay of Bengal Monsoon","m":[6,7,8,9],
     "a":5,"b":22,"c":78,"d":98,"f":0.82},
    {"n":"Arabian Sea Monsoon","m":[6,7,8],
     "a":5,"b":25,"c":50,"d":78,"f":0.80},
    {"n":"South Indian Winter","m":[6,7,8,9],
     "a":-60,"b":-35,"c":20,"d":115,"f":0.83},
    {"n":"South Pacific Winter","m":[6,7,8,9],
     "a":-60,"b":-35,"c":-180,"d":-60,"f":0.82},
    {"n":"NW Pacific Typhoons","m":[7,8,9,10],
     "a":10,"b":35,"c":110,"d":165,"f":0.85},
    {"n":"Caribbean Hurricane","m":[6,7,8,9,10],
     "a":10,"b":32,"c":-100,"d":-55,"f":0.88},
    {"n":"North Pacific Winter","m":[11,12,1,2,3],
     "a":35,"b":60,"c":145,"d":180,"f":0.85},
    {"n":"Cape Storms","m":[5,6,7,8,9],
     "a":-40,"b":-28,"c":10,"d":30,"f":0.84},
    {"n":"Red Sea Summer","m":[5,6,7,8,9],
     "a":12,"b":30,"c":32,"d":44,"f":0.93},
    {"n":"South China Sea NE Monsoon","m":[11,12,1,2],
     "a":5,"b":25,"c":105,"d":125,"f":0.88},
    {"n":"Barents / Norwegian Sea Winter","m":[10,11,12,1,2,3,4],
     "a":62,"b":75,"c":5,"d":45,"f":0.78},   # covers Vitino arctic route
]

SECA = [
    {"n":"Baltic ECA",          "a":53.5,"b":66.0,"c": 9.0,"d":30.0},
    {"n":"North Sea ECA",       "a":48.0,"b":62.0,"c":-5.0,"d": 9.0},
    {"n":"North American ECA",  "a":24.0,"b":50.0,"c":-140.0,"d":-50.0},
    {"n":"US Caribbean ECA",    "a":15.0,"b":31.0,"c":-98.0,"d":-60.0},
    {"n":"Chinese ECA",         "a":18.0,"b":41.0,"c":107.0,"d":125.0},
    {"n":"Mediterranean ECA",   "a":30.0,"b":46.0,"c":-6.0,"d":36.0},
]

RISK = [
    {"n":"Gulf of Aden / Somali Basin","t":"AWRP+LOH",
     "a":-5,"b":16,"c":45,"d":65,"aw":0.035,"lo":0.025},
    {"n":"Red Sea","t":"AWRP+LOH",
     "a":12,"b":30,"c":32,"d":44,"aw":0.050,"lo":0.035},
    {"n":"Persian Gulf","t":"AWRP+LOH",
     "a":22,"b":30,"c":48,"d":57,"aw":0.020,"lo":0.015},
    {"n":"Black Sea","t":"AWRP+LOH",
     "a":41,"b":47.5,"c":28,"d":37.5,"aw":0.120,"lo":0.080},
    {"n":"Korean Waters","t":"KR",
     "a":32,"b":42,"c":124,"d":132,"kr":3000},
    {"n":"Taiwan Strait","t":"KR",
     "a":22,"b":27,"c":118,"d":122.5,"kr":1500},
]

# ── GRAPH ─────────────────────────────────────────────────────────────────
ADJ: dict = {}
for _a, _b, _w in EDGES:
    ADJ.setdefault(_a, []).append((_b, _w))
    ADJ.setdefault(_b, []).append((_a, _w))

def haversine(a, b):
    R = 3440.065
    p1, p2 = math.radians(a[0]), math.radians(b[0])
    dp = p2 - p1
    dl = math.radians(b[1] - a[1])
    h = math.sin(dp/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 2 * R * math.asin(min(1.0, math.sqrt(h)))

def dijkstra(s, d):
    if s not in NODES or d not in NODES: return None
    if s == d: return {"total": 0.0, "legs": []}
    dist = {s: 0.0}; prev = {}
    h = [(0.0, s)]; done = set()
    while h:
        dd, u = heapq.heappop(h)
        if u in done: continue
        done.add(u)
        if u == d: break
        for v, w in ADJ.get(u, []):
            nd = dd + w
            if nd < dist.get(v, math.inf):
                dist[v] = nd; prev[v] = (u, w)
                heapq.heappush(h, (nd, v))
    if d not in dist: return None
    legs = []; cur = d
    while cur != s:
        p, w = prev[cur]; legs.append({"from": p, "to": cur, "nm": w}); cur = p
    legs.reverse()
    return {"total": dist[d], "legs": legs}

def in_box(lat, lon, z):
    return z["a"] <= lat <= z["b"] and z["c"] <= lon <= z["d"]

def hits(na, nb, z):
    mid = ((na[0]+nb[0])/2, (na[1]+nb[1])/2)
    return any(in_box(p[0], p[1], z) for p in (na, nb, mid))

def season_factor(na, nb, month):
    if not month: return 1.0, []
    fs, ns = [], []
    for r in SEASONAL:
        if month in r["m"] and hits(na, nb, r):
            fs.append(r["f"]); ns.append(r["n"])
    return (min(fs) if fs else 1.0), ns

def seca_names(na, nb):
    return [z["n"] for z in SECA if hits(na, nb, z)]

def risk_zones(na, nb):
    return [z for z in RISK if hits(na, nb, z)]

def get_bprice(port, fuel, bunker_ref):
    for b in bunker_ref:
        if b["port"] == port:
            return float(b.get(fuel, 0) or 0)
    return {"vlsfo": 590, "mgo": 930}.get(fuel, 590)

def get_port_costs(port, bunker_ref):
    for b in bunker_ref:
        if b["port"] == port:
            return (float(b.get("port_dues", 0) or 0),
                    float(b.get("canal_dues", 0) or 0),
                    float(b.get("agents", 0) or 0))
    return 18000, 0, 7000

# ── DEFAULT REFERENCES ────────────────────────────────────────────────────
DEFAULT_CARGOES = [
    {"code":"SFO","name":"Подсолнечное масло","group":"Soft oil","density":0.918,
     "heat_deg":0,"heat_mt":0,"parcel_min":1000,"parcel_max":35000,
     "tanks":"coated","notes":"FFA sensitive, avoid heating above 30°C"},
    {"code":"SBO","name":"Соевое масло","group":"Soft oil","density":0.920,
     "heat_deg":0,"heat_mt":0,"parcel_min":1000,"parcel_max":40000,
     "tanks":"coated","notes":"Standard parcel tanker grade"},
    {"code":"RSO","name":"Рапсовое масло","group":"Soft oil","density":0.914,
     "heat_deg":0,"heat_mt":0,"parcel_min":1000,"parcel_max":30000,
     "tanks":"coated","notes":"EU biodiesel feedstock"},
    {"code":"CPO","name":"Пальмовое масло сырое","group":"Palm","density":0.891,
     "heat_deg":50,"heat_mt":1.8,"parcel_min":3000,"parcel_max":50000,
     "tanks":"stainless or coated","notes":"Heating 50-55°C at sea and in port"},
    {"code":"RBDPO","name":"Пальмовое масло RBD","group":"Palm","density":0.891,
     "heat_deg":45,"heat_mt":1.5,"parcel_min":2000,"parcel_max":50000,
     "tanks":"stainless","notes":"Edible grade, strict FFA limit"},
    {"code":"POL","name":"Пальмовый олеин","group":"Palm fraction","density":0.910,
     "heat_deg":40,"heat_mt":1.2,"parcel_min":2000,"parcel_max":45000,
     "tanks":"stainless","notes":"Cloud point controlled"},
    {"code":"PST","name":"Пальмовый стеарин","group":"Palm fraction","density":0.880,
     "heat_deg":60,"heat_mt":2.4,"parcel_min":2000,"parcel_max":40000,
     "tanks":"stainless","notes":"High melt point, keep 60-65°C"},
    {"code":"PFAD","name":"Palm fatty acid distillate","group":"Palm fraction",
     "density":0.870,"heat_deg":60,"heat_mt":2.6,"parcel_min":2000,"parcel_max":40000,
     "tanks":"stainless","notes":"Corrosive, dedicated tanks"},
    {"code":"PKO","name":"Пальмоядровое масло","group":"Lauric","density":0.905,
     "heat_deg":35,"heat_mt":1.0,"parcel_min":1000,"parcel_max":25000,
     "tanks":"stainless","notes":"Lauric grade"},
    {"code":"CNO","name":"Кокосовое масло","group":"Lauric","density":0.908,
     "heat_deg":35,"heat_mt":1.0,"parcel_min":1000,"parcel_max":20000,
     "tanks":"stainless","notes":"Lauric grade"},
]

DEFAULT_VESSELS = [
    {"name":"MT Elaeis Trader","cls":"Palm tanker 20k","intake":19900,"speed":14.0,
     "cons_sea":26,"cons_mgo":19,"cons_port":3.0,"tank":1100,"port_hours":34,
     "load_hours":30,"disch_hours":38,"status":"owned","tc_rate":9200,"opex_day":7000,
     "value":22000000,"offhire_days":12,"notes":"Stainless tanks, 28 cargo tanks"},
    {"name":"MT Olea Star","cls":"Handy tanker 13k","intake":12800,"speed":13.0,
     "cons_sea":20,"cons_mgo":15,"cons_port":2.4,"tank":760,"port_hours":30,
     "load_hours":26,"disch_hours":34,"status":"owned","tc_rate":8400,"opex_day":6200,
     "value":16000000,"offhire_days":14,"notes":"14 tanks coated, IMO 2"},
    {"name":"MT Palm Express","cls":"MR tanker 25k","intake":24500,"speed":14.5,
     "cons_sea":30,"cons_mgo":22,"cons_port":3.5,"tank":1400,"port_hours":40,
     "load_hours":36,"disch_hours":44,"status":"chartered","tc_rate":16800,"opex_day":0,
     "value":32000000,"offhire_days":8,"notes":"TC-in, 22 stainless tanks"},
    {"name":"MT Solaris","cls":"MR tanker 37k","intake":36800,"speed":14.0,
     "cons_sea":36,"cons_mgo":26,"cons_port":4.0,"tank":1800,"port_hours":52,
     "load_hours":48,"disch_hours":56,"status":"owned","tc_rate":11200,"opex_day":8500,
     "value":44000000,"offhire_days":10,"notes":"33 stainless + epoxy tanks"},
    {"name":"MT Oleander","cls":"Chemical 8k","intake":7900,"speed":12.5,
     "cons_sea":16,"cons_mgo":12,"cons_port":2.0,"tank":560,"port_hours":26,
     "load_hours":22,"disch_hours":28,"status":"owned","tc_rate":7800,"opex_day":5800,
     "value":12000000,"offhire_days":18,"notes":"Multi-grade 18 tanks, IMO 2"},
]

# bunker prices for key edible-oil loading/discharging ports
DEFAULT_BUNKER = [
    {"port":"Rotterdam",     "vlsfo":585,"mgo":905,"port_dues":44000,"canal_dues":0,"agents":9000},
    {"port":"Amsterdam",     "vlsfo":588,"mgo":912,"port_dues":41000,"canal_dues":0,"agents":8500},
    {"port":"Antwerp",       "vlsfo":592,"mgo":918,"port_dues":40000,"canal_dues":0,"agents":8800},
    {"port":"Hamburg",       "vlsfo":598,"mgo":925,"port_dues":38000,"canal_dues":0,"agents":8200},
    {"port":"Le Havre",      "vlsfo":600,"mgo":930,"port_dues":39000,"canal_dues":0,"agents":8000},
    {"port":"Ghent",         "vlsfo":594,"mgo":920,"port_dues":36000,"canal_dues":0,"agents":8200},
    {"port":"St. Petersburg","vlsfo":620,"mgo":965,"port_dues":30000,"canal_dues":0,"agents":11000},
    {"port":"Gdansk",        "vlsfo":610,"mgo":955,"port_dues":28000,"canal_dues":0,"agents":9500},
    {"port":"Vitino",        "vlsfo":640,"mgo":990,"port_dues":22000,"canal_dues":0,"agents":14000},
    {"port":"Novorossiysk",  "vlsfo":610,"mgo":960,"port_dues":26000,"canal_dues":0,"agents":11000},
    {"port":"Taman",         "vlsfo":612,"mgo":962,"port_dues":24000,"canal_dues":0,"agents":10500},
    {"port":"Odessa",        "vlsfo":615,"mgo":965,"port_dues":27000,"canal_dues":0,"agents":12000},
    {"port":"Constanta",     "vlsfo":608,"mgo":955,"port_dues":25000,"canal_dues":0,"agents":9500},
    {"port":"Istanbul",      "vlsfo":605,"mgo":950,"port_dues":22000,"canal_dues":0,"agents":9000},
    {"port":"Piraeus",       "vlsfo":602,"mgo":945,"port_dues":24000,"canal_dues":0,"agents":8800},
    {"port":"Port Said",     "vlsfo":598,"mgo":940,"port_dues":18000,"canal_dues":180000,"agents":12000},
    {"port":"Alexandria",    "vlsfo":600,"mgo":940,"port_dues":21000,"canal_dues":0,"agents":8600},
    {"port":"Damietta",      "vlsfo":602,"mgo":942,"port_dues":20000,"canal_dues":0,"agents":8400},
    {"port":"Jeddah",        "vlsfo":612,"mgo":965,"port_dues":32000,"canal_dues":0,"agents":10000},
    {"port":"Fujairah",      "vlsfo":590,"mgo":935,"port_dues":30000,"canal_dues":0,"agents":9500},
    {"port":"Jebel Ali",     "vlsfo":592,"mgo":937,"port_dues":31000,"canal_dues":0,"agents":9500},
    {"port":"Karachi",       "vlsfo":618,"mgo":975,"port_dues":24000,"canal_dues":0,"agents":9000},
    {"port":"Port Qasim",    "vlsfo":620,"mgo":977,"port_dues":22000,"canal_dues":0,"agents":8800},
    {"port":"Mumbai",        "vlsfo":622,"mgo":980,"port_dues":28000,"canal_dues":0,"agents":9500},
    {"port":"Kandla",        "vlsfo":625,"mgo":985,"port_dues":23000,"canal_dues":0,"agents":8800},
    {"port":"Mundra",        "vlsfo":624,"mgo":984,"port_dues":22000,"canal_dues":0,"agents":8600},
    {"port":"New Mangalore", "vlsfo":626,"mgo":986,"port_dues":20000,"canal_dues":0,"agents":8400},
    {"port":"Chennai",       "vlsfo":620,"mgo":978,"port_dues":25000,"canal_dues":0,"agents":9000},
    {"port":"Krishnapatnam", "vlsfo":621,"mgo":979,"port_dues":19000,"canal_dues":0,"agents":8200},
    {"port":"Kakinada",      "vlsfo":622,"mgo":980,"port_dues":18000,"canal_dues":0,"agents":8000},
    {"port":"Kolkata",       "vlsfo":622,"mgo":980,"port_dues":26000,"canal_dues":0,"agents":9200},
    {"port":"Haldia",        "vlsfo":624,"mgo":982,"port_dues":24000,"canal_dues":0,"agents":8800},
    {"port":"Chittagong",    "vlsfo":628,"mgo":988,"port_dues":27000,"canal_dues":0,"agents":9800},
    {"port":"Colombo",       "vlsfo":610,"mgo":960,"port_dues":20000,"canal_dues":0,"agents":8000},
    {"port":"Singapore",     "vlsfo":578,"mgo":915,"port_dues":34000,"canal_dues":0,"agents":9000},
    {"port":"Port Klang",    "vlsfo":582,"mgo":920,"port_dues":28000,"canal_dues":0,"agents":8200},
    {"port":"Pasir Gudang",  "vlsfo":580,"mgo":918,"port_dues":27000,"canal_dues":0,"agents":8000},
    {"port":"Belawan",       "vlsfo":592,"mgo":930,"port_dues":24000,"canal_dues":0,"agents":7600},
    {"port":"Dumai",         "vlsfo":590,"mgo":928,"port_dues":22000,"canal_dues":0,"agents":7400},
    {"port":"Kuala Tanjung", "vlsfo":591,"mgo":929,"port_dues":21000,"canal_dues":0,"agents":7200},
    {"port":"Tanjung Priok", "vlsfo":588,"mgo":925,"port_dues":25000,"canal_dues":0,"agents":7800},
    {"port":"Bangkok",       "vlsfo":596,"mgo":934,"port_dues":26000,"canal_dues":0,"agents":8200},
    {"port":"Ho Chi Minh",   "vlsfo":594,"mgo":932,"port_dues":24000,"canal_dues":0,"agents":8000},
    {"port":"Hong Kong",     "vlsfo":600,"mgo":940,"port_dues":36000,"canal_dues":0,"agents":9500},
    {"port":"Shanghai",      "vlsfo":602,"mgo":942,"port_dues":35000,"canal_dues":0,"agents":9200},
    {"port":"Nansha",        "vlsfo":600,"mgo":940,"port_dues":33000,"canal_dues":0,"agents":9000},
    {"port":"Tianjin",       "vlsfo":608,"mgo":948,"port_dues":36000,"canal_dues":0,"agents":9500},
    {"port":"Qingdao",       "vlsfo":606,"mgo":946,"port_dues":34000,"canal_dues":0,"agents":9200},
    {"port":"Busan",         "vlsfo":596,"mgo":936,"port_dues":34000,"canal_dues":0,"agents":9000},
    {"port":"Kobe",          "vlsfo":610,"mgo":950,"port_dues":38000,"canal_dues":0,"agents":10000},
    {"port":"Yokohama",      "vlsfo":612,"mgo":952,"port_dues":39000,"canal_dues":0,"agents":10200},
    {"port":"New York",      "vlsfo":605,"mgo":945,"port_dues":48000,"canal_dues":0,"agents":12000},
    {"port":"Houston",       "vlsfo":595,"mgo":935,"port_dues":42000,"canal_dues":0,"agents":11000},
    {"port":"New Orleans",   "vlsfo":592,"mgo":932,"port_dues":40000,"canal_dues":0,"agents":10500},
    {"port":"Los Angeles",   "vlsfo":608,"mgo":948,"port_dues":50000,"canal_dues":0,"agents":12500},
    {"port":"Santos",        "vlsfo":618,"mgo":972,"port_dues":32000,"canal_dues":0,"agents":11000},
    {"port":"Paranagua",     "vlsfo":620,"mgo":974,"port_dues":30000,"canal_dues":0,"agents":10500},
    {"port":"Rosario",       "vlsfo":622,"mgo":978,"port_dues":28000,"canal_dues":0,"agents":12000},
    {"port":"Buenos Aires",  "vlsfo":620,"mgo":975,"port_dues":30000,"canal_dues":0,"agents":11500},
    {"port":"Las Palmas",    "vlsfo":588,"mgo":920,"port_dues":22000,"canal_dues":0,"agents":8000},
    {"port":"Dakar",         "vlsfo":606,"mgo":950,"port_dues":20000,"canal_dues":0,"agents":8500},
    {"port":"Mombasa",       "vlsfo":616,"mgo":968,"port_dues":24000,"canal_dues":0,"agents":9000},
    {"port":"Durban",        "vlsfo":612,"mgo":962,"port_dues":26000,"canal_dues":0,"agents":9200},
    {"port":"Cape Town",     "vlsfo":610,"mgo":960,"port_dues":25000,"canal_dues":0,"agents":9000},
    {"port":"Fremantle",     "vlsfo":610,"mgo":960,"port_dues":22000,"canal_dues":0,"agents":9000},
    {"port":"Sydney",        "vlsfo":618,"mgo":972,"port_dues":28000,"canal_dues":0,"agents":9500},
    {"port":"Panama",        "vlsfo":598,"mgo":938,"port_dues":30000,"canal_dues":240000,"agents":14000},
]

CONTRACT_TYPES = [
    "Voyage Charter", "Spot", "Time Charter In",
    "Time Charter Out", "CoA", "Lastage",
]
MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun",
               "Jul","Aug","Sep","Oct","Nov","Dec"]

# ── CALCULATION ENGINE ────────────────────────────────────────────────────
def calc_shipment(shp, vessel, bunker_ref):
    """
    Charterer perspective.
    Total cost = bunker (VLSFO + MGO + heating) + port/canal + hire/TC + risk.
    Cost per tonne = total_cost / quantity.
    TCE = (revenue − variable costs) / days.
    """
    load  = shp.get("load_port", "")
    disch = shp.get("disch_port", "")
    if not load or not disch or load not in NODES or disch not in NODES:
        return None
    r = dijkstra(load, disch)
    if not r or not r["legs"]: return None

    spd     = float(vessel.get("speed")      or 13)
    cons_s  = float(vessel.get("cons_sea")   or 26)
    cons_m  = float(vessel.get("cons_mgo")   or 19)
    cons_p  = float(vessel.get("cons_port")  or  3)
    tank    = float(vessel.get("tank")       or 1100)
    load_h  = float(vessel.get("load_hours") or 30)
    disch_h = float(vessel.get("disch_hours")or 36)
    value   = float(vessel.get("value")      or 22_000_000)
    month   = int(shp.get("month") or 0)
    qty     = float(shp.get("quantity")      or 0)
    fr_rate = float(shp.get("freight_rate")  or 0)
    heat_mt = float(shp.get("heat_mt_day")   or 0)
    contract= shp.get("contract_type", "Voyage Charter")
    tc_rate = float(vessel.get("tc_rate")    or 14500)
    opex    = float(vessel.get("opex_day")   or 7000)
    spot_r  = float(shp.get("spot_rate")     or tc_rate)

    rob = tank
    ifo_sea = mgo_sea = days_sea = sea_nm = 0.0
    awrp = loh_v = kr = 0.0
    seca_set: set = set()
    risk_set: list = []
    alerts: list = []
    sfmin = 1.0

    for L in r["legs"]:
        if L["from"] not in NODES or L["to"] not in NODES: continue
        na, nb = NODES[L["from"]], NODES[L["to"]]
        sf, _ = season_factor(na, nb, month)
        sfmin = min(sfmin, sf)
        eff = spd * sf
        dz  = L["nm"] / (eff * 24) if eff > 0 else 0
        sz  = seca_names(na, nb)
        if sz:
            need = cons_m * dz
            if rob < need * 1.05:
                alerts.append({"port": L["from"],
                    "mt": round(need * 1.05 - rob, 1), "zone": sz[0]})
            mgo_sea += need; rob = max(0.0, rob - need); seca_set.update(sz)
        else:
            need = cons_s * dz; ifo_sea += need; rob = max(0.0, rob - need)
        days_sea += dz; sea_nm += L["nm"]
        for rz in risk_zones(na, nb):
            risk_set.append(rz)
            awrp  += rz.get("aw", 0) * value / 100
            loh_v += rz.get("lo", 0) * value / 100
            kr    += rz.get("kr", 0) * dz

    port_days = (load_h + disch_h) / 24
    ifo_sea  += cons_p * port_days
    total_days = days_sea + port_days

    vlsfo_p = get_bprice(load, "vlsfo", bunker_ref)
    mgo_p   = get_bprice(load, "mgo",   bunker_ref)
    heat_cost  = heat_mt * total_days * qty / 1000 * vlsfo_p
    bunk_cost  = ifo_sea * vlsfo_p + mgo_sea * mgo_p + heat_cost

    pd_l, cd_l, ag_l = get_port_costs(load,  bunker_ref)
    pd_d, cd_d, ag_d = get_port_costs(disch, bunker_ref)
    port_cost  = pd_l + pd_d + ag_l + ag_d + cd_l + cd_d
    risk_cost  = awrp + loh_v + kr

    if contract == "Time Charter In":
        hire_cost = tc_rate * total_days; opex_cost = 0.0
    elif contract == "Spot":
        hire_cost = spot_r * total_days;  opex_cost = 0.0
    elif contract in ("CoA", "Lastage"):
        hire_cost = tc_rate * total_days * 0.95; opex_cost = 0.0
    elif contract == "Time Charter Out":
        hire_cost = tc_rate * total_days; opex_cost = 0.0
    else:   # Voyage Charter
        hire_cost = 0.0; opex_cost = opex * total_days

    total_cost  = hire_cost + bunk_cost + port_cost + risk_cost + opex_cost
    revenue     = qty * fr_rate
    margin      = revenue - total_cost
    tce         = (revenue - bunk_cost - port_cost - risk_cost) / total_days \
                  if total_days > 0 else 0
    cpt         = total_cost / qty if qty > 0 else 0   # cost per tonne

    return {
        "load": load, "disch": disch, "nm": round(sea_nm),
        "days_sea":   round(days_sea, 2),
        "port_days":  round(port_days, 2),
        "total_days": round(total_days, 2),
        "ifo_mt":     round(ifo_sea, 1),
        "lsmgo_mt":   round(mgo_sea, 1),
        "heat_cost":  round(heat_cost),
        "bunker_cost":round(bunk_cost),
        "port_cost":  round(port_cost),
        "hire_cost":  round(hire_cost),
        "opex_cost":  round(opex_cost),
        "awrp":       round(awrp),
        "loh":        round(loh_v),
        "kr":         round(kr),
        "risk_cost":  round(risk_cost),
        "total_cost": round(total_cost),
        "revenue":    round(revenue),
        "margin":     round(margin),
        "tce":        round(tce),
        "cpt":        round(cpt, 2),
        "sfactor":    round(sfmin, 3),
        "seca":       list(seca_set),
        "risk_zones": [rz["n"] for rz in risk_set],
        "bunker_alerts": alerts,
        "contract":   contract,
        "vlsfo_price":vlsfo_p, "mgo_price": mgo_p,
        "legs":       r["legs"],
    }

# ── EXPORT HELPERS ────────────────────────────────────────────────────────
def build_xlsx(calc_data, vessels, bunker):
    wb = openpyxl.Workbook()
    INK = "15202B"; BLU = "38648A"; WSH = "DFE7EE"

    def hdr(ws, row, vals):
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=row, column=ci, value=v)
            c.font = Font(bold=True, color="FFFFFF", size=10)
            c.fill = PatternFill("solid", fgColor=BLU)
            c.alignment = Alignment(horizontal="center")
            c.border = Border(
                left=Side(style="thin", color=INK),
                right=Side(style="thin", color=INK),
                top=Side(style="thin", color=INK),
                bottom=Side(style="thin", color=INK))

    def row(ws, ri, vals):
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=v)
            c.font = Font(size=10)
            if ri % 2 == 0:
                c.fill = PatternFill("solid", fgColor=WSH)
            c.border = Border(
                left=Side(style="thin", color="C3CFD9"),
                right=Side(style="thin", color="C3CFD9"),
                top=Side(style="thin", color="C3CFD9"),
                bottom=Side(style="thin", color="C3CFD9"))

    # Sheet 1 – Shipments
    ws = wb.active; ws.title = "Отправки"
    cols = ["ID","Линия","Груз","Судно","Контракт","Порт A","Порт Б",
            "Кол-во mt","Дни","VLSFO mt","MGO mt","Выручка $","Бункер $",
            "Порт $","Найм $","Риск $","Итого $","Маржа $",
            "TCE $/d","Затр/т $","SECA","Риск-зоны"]
    hdr(ws, 1, cols)
    for ci in range(1, len(cols)+1):
        ws.column_dimensions[get_column_letter(ci)].width = 16
    for ri, r_ in enumerate((calc_data or {}).get("shipments", []), 2):
        row(ws, ri, [
            r_.get("shipment_id",""), r_.get("line_id",""),
            r_.get("cargo",""), r_.get("vessel",""), r_.get("contract",""),
            r_["load"], r_["disch"],
            r_.get("qty", 0), r_["total_days"],
            r_["ifo_mt"], r_["lsmgo_mt"],
            r_["revenue"], r_["bunker_cost"], r_["port_cost"],
            r_["hire_cost"], r_["risk_cost"], r_["total_cost"],
            r_["margin"], r_["tce"], r_["cpt"],
            "; ".join(r_.get("seca", [])),
            "; ".join(r_.get("risk_zones", [])),
        ])

    # Sheet 2 – Monthly budget
    ws2 = wb.create_sheet("Бюджет")
    bcols = ["Месяц","Выручка $","Бункер $","Найм $","Риск $",
             "Маржа $","Объём mt","Судо-дни"]
    hdr(ws2, 1, bcols)
    for ci in range(1, len(bcols)+1):
        ws2.column_dimensions[get_column_letter(ci)].width = 18
    rev_m=[0.]*12; cost_m=[0.]*12; bunk_m=[0.]*12
    hire_m=[0.]*12; risk_m=[0.]*12; qty_m=[0.]*12; days_m=[0.]*12
    for r_ in (calc_data or {}).get("shipments", []):
        m = int(r_.get("month") or 0)
        idxs = [m-1] if 1 <= m <= 12 else list(range(12))
        w = 1.0 if len(idxs)==1 else 1./12
        for i in idxs:
            rev_m[i]  += r_["revenue"]    * w
            bunk_m[i] += r_["bunker_cost"]* w
            hire_m[i] += r_["hire_cost"]  * w
            risk_m[i] += r_["risk_cost"]  * w
            cost_m[i] += r_["total_cost"] * w
            qty_m[i]  += r_.get("qty", 0) * w
            days_m[i] += r_["total_days"] * w
    for i in range(12):
        row(ws2, i+2, [MONTH_NAMES[i], round(rev_m[i]), round(bunk_m[i]),
            round(hire_m[i]), round(risk_m[i]),
            round(rev_m[i]-cost_m[i]), round(qty_m[i]), round(days_m[i],1)])
    row(ws2, 14, ["ИТОГО", round(sum(rev_m)), round(sum(bunk_m)),
        round(sum(hire_m)), round(sum(risk_m)),
        round(sum(rev_m)-sum(cost_m)), round(sum(qty_m)), round(sum(days_m),1)])
    for ci in range(1, 9):
        c = ws2.cell(14, ci)
        c.font = Font(bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=INK)

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()

def build_pdf(calc_data):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
        leftMargin=1*cm, rightMargin=1*cm,
        topMargin=1.5*cm, bottomMargin=1.5*cm)
    ink_c  = rlc.HexColor("#15202B")
    blue_c = rlc.HexColor("#38648A")
    wash_c = rlc.HexColor("#E2EDF4")
    white  = rlc.white
    styles = getSampleStyleSheet()
    title_s = ParagraphStyle("T", parent=styles["Normal"],
        fontSize=13, textColor=ink_c, fontName="Helvetica-Bold", spaceAfter=6)
    sub_s = ParagraphStyle("S", parent=styles["Normal"],
        fontSize=9, textColor=rlc.HexColor("#566877"),
        fontName="Helvetica", spaceAfter=4)
    elems = [
        Paragraph("EDIBLE OILS CHARTERER'S PLANNER – СВОДНЫЙ ОТЧЁТ", title_s),
        HRFlowable(width="100%", thickness=1.5, color=ink_c),
        Spacer(1, 0.3*cm),
    ]
    shp_res = (calc_data or {}).get("shipments", [])
    if not shp_res:
        elems.append(Paragraph("Нет рассчитанных отправок.", sub_s))
    else:
        total_rev  = sum(r["revenue"]    for r in shp_res)
        total_cost = sum(r["total_cost"] for r in shp_res)
        total_qty  = sum(r.get("qty", 0) for r in shp_res)
        avg_cpt    = total_cost / total_qty if total_qty else 0
        summ_data  = [
            ["Показатель", "Значение"],
            ["Отправок",              str(len(shp_res))],
            ["Суммарная выручка",     f"${total_rev:,.0f}"],
            ["Суммарные затраты",     f"${total_cost:,.0f}"],
            ["Маржа",                 f"${total_rev-total_cost:,.0f}"],
            ["Объём",                 f"{total_qty:,.0f} mt"],
            ["Средние затраты / тн",  f"${avg_cpt:,.2f}"],
        ]
        ts = TableStyle([
            ("BACKGROUND",(0,0),(1,0),blue_c),("TEXTCOLOR",(0,0),(1,0),white),
            ("FONTNAME",(0,0),(1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),9),
            ("GRID",(0,0),(-1,-1),0.5,ink_c),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[white,wash_c]),
            ("FONTNAME",(0,1),(-1,-1),"Helvetica"),
        ])
        t = Table(summ_data, colWidths=[9*cm, 9*cm])
        t.setStyle(ts)
        elems.extend([t, Spacer(1, 0.4*cm)])
        elems.append(Paragraph("Детализация отправок", title_s))
        leg_h = ["ID","Маршрут","Судно","mt","Дни",
                 "$/тн","Выручка","Затраты","Маржа","TCE $/d"]
        leg_d = [leg_h] + [[
            r.get("shipment_id",""),
            f"{r['load'][:12]}→{r['disch'][:12]}",
            r.get("vessel","")[:12],
            f"{r.get('qty',0):,.0f}", f"{r['total_days']:.1f}",
            f"${r['cpt']:,.2f}",
            f"${r['revenue']:,.0f}", f"${r['total_cost']:,.0f}",
            f"${r['margin']:,.0f}", f"${r['tce']:,.0f}",
        ] for r in shp_res]
        widths = [2.0,5.0,3.2,2.2,1.5,2.0,2.5,2.5,2.5,2.2]
        tleg = Table(leg_d, colWidths=[w*cm for w in widths])
        tleg.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),blue_c),("TEXTCOLOR",(0,0),(-1,0),white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),7.5),
            ("GRID",(0,0),(-1,-1),0.4,ink_c),
            ("FONTNAME",(0,1),(-1,-1),"Helvetica"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[white,wash_c]),
            ("ALIGN",(3,0),(-1,-1),"CENTER"),
        ]))
        elems.append(tleg)
    doc.build(elems); buf.seek(0)
    return buf.getvalue()

# ── DASH APP ──────────────────────────────────────────────────────────────
app = Dash(__name__, suppress_callback_exceptions=True)
app.title = "Edible Oils Planner"

port_list   = sorted(NODES.keys())
port_opts   = [{"label": k, "value": k} for k in port_list]
vessel_opts = [{"label": v["name"], "value": v["name"]} for v in DEFAULT_VESSELS]
ctract_opts = [{"label": t, "value": t} for t in CONTRACT_TYPES]
month_opts  = [{"label": "Любой", "value": 0}] + \
              [{"label": MONTH_NAMES[i], "value": i+1} for i in range(12)]

# map background layers (Codespaces: no ngrok, served locally)
seca_rects = [
    dl.Rectangle(bounds=[[z["a"],z["c"]],[z["b"],z["d"]]],
        color=C["green"], weight=1.5, fill=True,
        fillColor=C["green_l"], fillOpacity=0.18,
        children=dl.Tooltip(f"SECA: {z['n']}"))
    for z in SECA
]
risk_rects = [
    dl.Rectangle(bounds=[[z["a"],z["c"]],[z["b"],z["d"]]],
        color=C["red"], weight=1.5, dashArray="5 4",
        fill=True, fillColor=C["amber_l"], fillOpacity=0.20,
        children=dl.Tooltip(f"{z['t']}: {z['n']}"))
    for z in RISK
]
junction_dots = [
    dl.CircleMarker(center=[v[0],v[1]], radius=4,
        color=C["ink"], weight=1, fillColor=C["blue_l"], fillOpacity=0.85,
        children=dl.Tooltip(k))
    for k, v in NODES.items() if v[2] == "j"
]

def tab_style(sel=False):
    base = {"fontWeight": "600", "fontSize": "12px"}
    if sel:
        return {**base, "fontWeight": "800", "color": C["blue"],
                "borderTop": f"2px solid {C['blue']}"}
    return base

app.layout = html.Div([
    # top bar
    html.Div([
        html.Span("EDIBLE OILS FLEET PLANNER",
            style={"fontWeight":"800","fontSize":"14px",
                   "letterSpacing":"0.07em","color":"#fff"}),
        html.Span("  ·  charterer perspective · shipments · lines · budget · spot",
            style={"fontSize":"11px","color":C["blue_l"],"marginLeft":"8px"}),
    ], style={"padding":"9px 16px","background":C["ink"],
              "borderBottom":f"2px solid {C['blue']}","display":"flex",
              "alignItems":"center"}),

    dcc.Tabs(id="main-tabs", value="tab-shipments",
        colors={"border":C["ink"],"primary":C["blue"],"background":C["bg"]},
        style={"fontFamily": F},
        children=[
            dcc.Tab(label="Справочники",   value="tab-refs",
                style=tab_style(), selected_style=tab_style(True)),
            dcc.Tab(label="Отправки",      value="tab-shipments",
                style=tab_style(), selected_style=tab_style(True)),
            dcc.Tab(label="Линии",         value="tab-lines",
                style=tab_style(), selected_style=tab_style(True)),
            dcc.Tab(label="Флот / покрытие",value="tab-fleet",
                style=tab_style(), selected_style=tab_style(True)),
            dcc.Tab(label="Годовой план",  value="tab-plan",
                style=tab_style(), selected_style=tab_style(True)),
            dcc.Tab(label="Спот / рынок",  value="tab-spot",
                style=tab_style(), selected_style=tab_style(True)),
            dcc.Tab(label="Карта",         value="tab-map",
                style=tab_style(), selected_style=tab_style(True)),
            dcc.Tab(label="Экспорт",       value="tab-export",
                style=tab_style(), selected_style=tab_style(True)),
        ]),

    html.Div(id="tab-content",
        style={"padding":"12px 14px","overflowY":"auto",
               "flex":"1","background":C["bg"]}),

    # stores
    dcc.Store(id="s-cargoes",   data=DEFAULT_CARGOES),
    dcc.Store(id="s-bunker",    data=DEFAULT_BUNKER),
    dcc.Store(id="s-vessels",   data=DEFAULT_VESSELS),
    dcc.Store(id="s-shipments", data=[]),
    dcc.Store(id="s-calc",      data={}),

    dcc.Download(id="dl-xlsx"),
    dcc.Download(id="dl-pdf"),
    dcc.Download(id="dl-json"),
], style={"display":"flex","flexDirection":"column","height":"100vh",
          "fontFamily":F,"fontSize":"13px","color":C["ink"],"background":C["bg"]})

# ── TAB RENDERER ──────────────────────────────────────────────────────────
@app.callback(
    Output("tab-content","children"),
    Input("main-tabs","value"),
    State("s-cargoes","data"),
    State("s-bunker","data"),
    State("s-vessels","data"),
    State("s-shipments","data"),
    State("s-calc","data"),
)
def render_tab(tab, cargoes, bunker, vessels, shipments, calc):

    # ── СПРАВОЧНИКИ ───────────────────────────────────────────────────────
    if tab == "tab-refs":
        cargo_cols = [
            {"name":"Код",     "id":"code",       "editable":True},
            {"name":"Название","id":"name",        "editable":True},
            {"name":"Группа",  "id":"group",       "editable":True},
            {"name":"Плотность","id":"density",    "editable":True,"type":"numeric"},
            {"name":"Подогрев °C","id":"heat_deg", "editable":True,"type":"numeric"},
            {"name":"Топл. подогр. mt/d","id":"heat_mt","editable":True,"type":"numeric"},
            {"name":"Мин. партия","id":"parcel_min","editable":True,"type":"numeric"},
            {"name":"Макс. партия","id":"parcel_max","editable":True,"type":"numeric"},
            {"name":"Танки",   "id":"tanks",       "editable":True},
            {"name":"Заметки", "id":"notes",       "editable":True},
        ]
        bunker_cols = [
            {"name":"Порт",       "id":"port",       "editable":True},
            {"name":"VLSFO $/mt", "id":"vlsfo",      "editable":True,"type":"numeric"},
            {"name":"MGO $/mt",   "id":"mgo",        "editable":True,"type":"numeric"},
            {"name":"Port dues $","id":"port_dues",  "editable":True,"type":"numeric"},
            {"name":"Canal dues $","id":"canal_dues","editable":True,"type":"numeric"},
            {"name":"Agents $",   "id":"agents",     "editable":True,"type":"numeric"},
        ]
        vessel_cols = [
            {"name":"Судно",      "id":"name",        "editable":True},
            {"name":"Класс",      "id":"cls",         "editable":True},
            {"name":"Intake mt",  "id":"intake",      "editable":True,"type":"numeric"},
            {"name":"Скор. уз",   "id":"speed",       "editable":True,"type":"numeric"},
            {"name":"VLSFO/d",    "id":"cons_sea",    "editable":True,"type":"numeric"},
            {"name":"MGO/d SECA", "id":"cons_mgo",    "editable":True,"type":"numeric"},
            {"name":"Порт/d",     "id":"cons_port",   "editable":True,"type":"numeric"},
            {"name":"Бункер тнк", "id":"tank",        "editable":True,"type":"numeric"},
            {"name":"Порт ч (груз)","id":"load_hours","editable":True,"type":"numeric"},
            {"name":"Порт ч (выгр)","id":"disch_hours","editable":True,"type":"numeric"},
            {"name":"TC rate $/d","id":"tc_rate",     "editable":True,"type":"numeric"},
            {"name":"Opex $/d",   "id":"opex_day",    "editable":True,"type":"numeric"},
            {"name":"Стоим. $",   "id":"value",       "editable":True,"type":"numeric"},
            {"name":"Оффхайр дн.","id":"offhire_days","editable":True,"type":"numeric"},
            {"name":"Статус",     "id":"status",      "editable":True},
            {"name":"Заметки",    "id":"notes",       "editable":True},
        ]
        return html.Div([
            card([sec("Справочник грузов"),
                  html.P("Редактируйте прямо в таблице. Изменения применяются немедленно.",
                     style={"fontSize":"11px","color":C["muted"],"marginBottom":"8px"}),
                  dash_table.DataTable(id="tbl-cargoes", data=cargoes,
                      columns=cargo_cols, editable=True, row_deletable=True, **TBL)]),
            card([sec("Бункерные цены по портам"),
                  dash_table.DataTable(id="tbl-bunker", data=bunker,
                      columns=bunker_cols, editable=True, row_deletable=True, **TBL)]),
            card([sec("Справочник судов"),
                  dash_table.DataTable(id="tbl-vessels", data=vessels,
                      columns=vessel_cols, editable=True, row_deletable=True, **TBL)]),
        ])

    # ── ОТПРАВКИ ──────────────────────────────────────────────────────────
    elif tab == "tab-shipments":
        c_opts = [{"label":f"{c['code']} – {c['name']}","value":c["code"]}
                  for c in (cargoes or DEFAULT_CARGOES)]
        v_opts = [{"label":v["name"],"value":v["name"]}
                  for v in (vessels or DEFAULT_VESSELS)]
        shp_cols = [
            {"name":"ID",           "id":"id",            "editable":True},
            {"name":"Линия",        "id":"line_id",       "editable":True},
            {"name":"Груз",         "id":"cargo_code",    "editable":True,
             "presentation":"dropdown"},
            {"name":"Порт погрузки","id":"load_port",     "editable":True,
             "presentation":"dropdown"},
            {"name":"Порт выгрузки","id":"disch_port",    "editable":True,
             "presentation":"dropdown"},
            {"name":"Кол-во mt",    "id":"quantity",      "editable":True,"type":"numeric"},
            {"name":"Фрахт $/mt",   "id":"freight_rate",  "editable":True,"type":"numeric"},
            {"name":"Судно",        "id":"vessel_id",     "editable":True,
             "presentation":"dropdown"},
            {"name":"Контракт",     "id":"contract_type", "editable":True,
             "presentation":"dropdown"},
            {"name":"Спот $/d",     "id":"spot_rate",     "editable":True,"type":"numeric"},
            {"name":"Месяц",        "id":"month",         "editable":True,"type":"numeric"},
            {"name":"Период нач.",  "id":"period_start",  "editable":True},
            {"name":"Период кон.",  "id":"period_end",    "editable":True},
            {"name":"Плотность",    "id":"density",       "editable":True,"type":"numeric"},
            {"name":"Подогрев mt/d","id":"heat_mt_day",   "editable":True,"type":"numeric"},
            {"name":"Статус",       "id":"status",        "editable":True},
        ]
        dropdowns = {
            "cargo_code":    {"options": c_opts},
            "load_port":     {"options": port_opts},
            "disch_port":    {"options": port_opts},
            "vessel_id":     {"options": v_opts},
            "contract_type": {"options": ctract_opts},
        }
        seed = shipments if shipments else [
            {"id":f"S{i+1:03d}","line_id":"L1","cargo_code":"CPO",
             "load_port":"Pasir Gudang","disch_port":"Rotterdam",
             "quantity":20000,"freight_rate":85,
             "vessel_id":DEFAULT_VESSELS[0]["name"],
             "contract_type":"Voyage Charter","spot_rate":14000,
             "month":0,"period_start":"","period_end":"",
             "density":0.891,"heat_mt_day":1.8,"status":"planned"}
            for i in range(4)
        ]
        return html.Div([
            card([sec("Грузовые отправки"),
                  html.P("Отправка — базовая единица планирования. "
                         "Привязывается к линии, судну и типу контракта.",
                     style={"fontSize":"11px","color":C["muted"],"marginBottom":"8px"}),
                  html.Div([
                      html.Button("Рассчитать все отправки", id="btn-calc-all",
                          style={**btn,"width":"auto","marginRight":"8px","marginTop":"0"}),
                      html.Button("Добавить строку", id="btn-add-shp",
                          style={**btng,"width":"auto","marginTop":"0"}),
                  ], style={"display":"flex","marginBottom":"10px"}),
                  dash_table.DataTable(id="tbl-shipments", data=seed,
                      columns=shp_cols, editable=True, row_deletable=True,
                      dropdown=dropdowns, **TBL),
                  html.Div(id="shp-results")]),
        ])

    # ── ЛИНИИ ─────────────────────────────────────────────────────────────
    elif tab == "tab-lines":
        shp_res = (calc or {}).get("shipments", [])
        if not shp_res:
            return card([sec("Линии"), html.P("Сначала рассчитайте отправки.",
                style={"color":C["muted"]})])
        by_line: dict = {}
        for r in shp_res:
            by_line.setdefault(r.get("line_id","—"), []).append(r)

        cards = []
        for lid, items in sorted(by_line.items()):
            tot_rev  = sum(x["revenue"]    for x in items)
            tot_cost = sum(x["total_cost"] for x in items)
            tot_qty  = sum(x.get("qty",0)  for x in items)
            tot_days = sum(x["total_days"] for x in items)
            avg_tce  = sum(x["tce"] for x in items) / len(items)
            avg_cpt  = tot_cost / tot_qty if tot_qty else 0

            monthly = [0.0]*12
            for x in items:
                m = int(x.get("month") or 0)
                idxs = [m-1] if 1<=m<=12 else list(range(12))
                w = 1. if len(idxs)==1 else 1./12
                for i in idxs: monthly[i] += x["revenue"] * w

            fig = go.Figure(go.Bar(x=MONTH_NAMES, y=monthly,
                marker_color=C["blue_l"],
                marker_line_color=C["ink"], marker_line_width=1))
            fig.update_layout(height=140,
                margin=dict(l=30,r=10,t=8,b=30),
                paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
                font=dict(family=F, size=10, color=C["ink"]),
                showlegend=False,
                xaxis=dict(showgrid=False),
                yaxis=dict(gridcolor=C["ink_l"]))

            tbl = dash_table.DataTable(
                data=[{
                    "ID": x.get("shipment_id",""),
                    "Судно": x.get("vessel",""),
                    "Контракт": x.get("contract",""),
                    "От": x["load"],"До": x["disch"],
                    "mt": f"{x.get('qty',0):,.0f}",
                    "Дни": x["total_days"],
                    "$/тн": f"{x['cpt']:,.2f}",
                    "Выручка": money(x["revenue"]),
                    "Затраты": money(x["total_cost"]),
                    "Маржа": money(x["margin"]),
                    "TCE $/d": f"{x['tce']:,.0f}",
                } for x in items],
                columns=[{"name":c,"id":c} for c in
                    ["ID","Судно","Контракт","От","До","mt","Дни",
                     "$/тн","Выручка","Затраты","Маржа","TCE $/d"]],
                **TBL)

            cards.append(card([
                html.Div([
                    html.Div(f"ЛИНИЯ {lid}",
                        style={"fontSize":"13px","fontWeight":"800","color":C["ink"]}),
                    html.Div(f"{len(items)} отправок · {tot_qty:,.0f} mt · "
                             f"{tot_days:.0f} судо-дней",
                        style={"fontSize":"11px","color":C["muted"]}),
                ], style={"marginBottom":"10px"}),
                html.Div([
                    kpi("Выручка",     money(tot_rev),    f"{tot_qty:,.0f} mt"),
                    kpi("Затраты",     money(tot_cost),   "всего"),
                    kpi("Маржа",       money(tot_rev-tot_cost),"",
                        C["green"] if tot_rev >= tot_cost else C["red"]),
                    kpi("Ср. TCE",     f"{avg_tce:,.0f} $/d",""),
                    kpi("Затр./тн",    f"${avg_cpt:,.2f}",""),
                ], style={"display":"flex","gap":"8px","flexWrap":"wrap",
                          "marginBottom":"12px"}),
                dcc.Graph(figure=fig, config={"displayModeBar":False}),
                html.Div(style={"height":"8px"}),
                tbl,
            ]))
        return html.Div(cards)

    # ── ФЛОТ / ПОКРЫТИЕ ───────────────────────────────────────────────────
    elif tab == "tab-fleet":
        shp_res = (calc or {}).get("shipments", [])
        if not shp_res:
            return card([sec("Покрытие флотом"),
                html.P("Сначала рассчитайте отправки.", style={"color":C["muted"]})])

        usage: dict = {}
        for v in (vessels or DEFAULT_VESSELS):
            usage[v["name"]] = {
                "cls": v.get("cls",""), "status": v.get("status",""),
                "avail": 365 - float(v.get("offhire_days") or 0),
                "used": 0., "n": 0, "rev": 0., "cost": 0., "margin": 0.,
                "tc": float(v.get("tc_rate") or 0),
                "opex": float(v.get("opex_day") or 0),
            }
        for r in shp_res:
            vn = r.get("vessel","")
            if vn in usage:
                usage[vn]["used"]   += r["total_days"]
                usage[vn]["n"]      += 1
                usage[vn]["rev"]    += r["revenue"]
                usage[vn]["cost"]   += r["total_cost"]
                usage[vn]["margin"] += r["margin"]

        names_v = list(usage.keys())
        used_v  = [usage[n]["used"] for n in names_v]
        idle_v  = [max(0., usage[n]["avail"]-usage[n]["used"]) for n in names_v]

        fig_u = go.Figure()
        fig_u.add_trace(go.Bar(name="Занято", x=names_v, y=used_v,
            marker_color=C["blue"], marker_line_color=C["ink"], marker_line_width=1))
        fig_u.add_trace(go.Bar(name="Простой", x=names_v, y=idle_v,
            marker_color=C["ink_l"], marker_line_color=C["ink"], marker_line_width=1))
        fig_u.update_layout(barmode="stack", height=240,
            margin=dict(l=40,r=10,t=10,b=60),
            paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
            font=dict(family=F, size=10, color=C["ink"]),
            legend=dict(orientation="h", y=-0.3),
            xaxis=dict(showgrid=False, tickangle=-30),
            yaxis=dict(title="дни", gridcolor=C["ink_l"]))

        rows = []
        for vn, u in usage.items():
            idle = max(0., u["avail"] - u["used"])
            pct  = 100 * u["used"] / u["avail"] if u["avail"] > 0 else 0
            idle_tc = idle * u["tc"] if u["status"] == "owned" else 0
            period_cost = u["avail"] * (u["tc"] if u["status"]=="chartered" else u["opex"])
            rows.append({
                "Судно":   vn, "Класс": u["cls"], "Статус": u["status"],
                "Доступно дн.": round(u["avail"],1),
                "Занято дн.":   round(u["used"],1),
                "Простой дн.":  round(idle,1),
                "Загрузка %":   round(pct,1),
                "Отправок":     u["n"],
                "Выручка $":    f"{u['rev']:,.0f}",
                "Период. затр.$": f"{period_cost:,.0f}",
                "Маржа $":      f"{u['margin']:,.0f}",
                "TC-out простой $": f"{idle_tc:,.0f}",
            })

        return html.Div([
            card([sec("Покрытие флотом — загрузка судов"),
                  dcc.Graph(figure=fig_u, config={"displayModeBar":False})]),
            card([dash_table.DataTable(data=rows,
                columns=[{"name":c,"id":c} for c in rows[0].keys()],
                **TBL)]),
        ])

    # ── ГОДОВОЙ ПЛАН ──────────────────────────────────────────────────────
    elif tab == "tab-plan":
        shp_res = (calc or {}).get("shipments", [])
        if not shp_res:
            return card([sec("Годовой план"),
                html.P("Сначала рассчитайте отправки.", style={"color":C["muted"]})])

        rev_m=[0.]*12; cost_m=[0.]*12; bunk_m=[0.]*12
        hire_m=[0.]*12; risk_m=[0.]*12; qty_m=[0.]*12; days_m=[0.]*12
        for r in shp_res:
            m = int(r.get("month") or 0)
            idxs = [m-1] if 1<=m<=12 else list(range(12))
            w = 1. if len(idxs)==1 else 1./12
            for i in idxs:
                rev_m[i]  += r["revenue"]    * w
                bunk_m[i] += r["bunker_cost"]* w
                hire_m[i] += r["hire_cost"]  * w
                risk_m[i] += r["risk_cost"]  * w
                cost_m[i] += r["total_cost"] * w
                qty_m[i]  += r.get("qty",0)  * w
                days_m[i] += r["total_days"] * w

        total_rev  = sum(rev_m)
        total_cost = sum(cost_m)
        total_qty  = sum(qty_m)
        avg_cpt    = total_cost / total_qty if total_qty else 0

        # Waterfall
        wf = go.Figure(go.Waterfall(
            orientation="v",
            measure=["absolute","relative","relative","relative","relative","total"],
            x=["Выручка","Бункер","Найм/Opex","Порт/Канал","Риск","Маржа"],
            y=[total_rev,
               -sum(r["bunker_cost"] for r in shp_res),
               -sum(r["hire_cost"]+r.get("opex_cost",0) for r in shp_res),
               -sum(r["port_cost"]   for r in shp_res),
               -sum(r["risk_cost"]   for r in shp_res),
               total_rev - total_cost],
            connector={"line":{"color":C["ink"],"width":1}},
            increasing={"marker":{"color":C["green_l"],"line":{"color":C["ink"],"width":1}}},
            decreasing={"marker":{"color":C["red_l"], "line":{"color":C["ink"],"width":1}}},
            totals={"marker":{"color":C["blue_l"],    "line":{"color":C["ink"],"width":1}}},
        ))
        wf.update_layout(height=280, margin=dict(l=40,r=10,t=20,b=30),
            paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
            font=dict(family=F, size=10, color=C["ink"]),
            yaxis=dict(gridcolor=C["ink_l"], tickprefix="$"),
            xaxis=dict(showgrid=False))

        # Monthly bar
        mb = go.Figure()
        mb.add_trace(go.Bar(name="Выручка", x=MONTH_NAMES, y=rev_m,
            marker_color=C["blue_l"],marker_line_color=C["ink"],marker_line_width=1))
        mb.add_trace(go.Bar(name="Затраты", x=MONTH_NAMES, y=cost_m,
            marker_color=C["red_l"], marker_line_color=C["ink"],marker_line_width=1))
        mb.add_trace(go.Scatter(name="Маржа", x=MONTH_NAMES,
            y=[r-c for r,c in zip(rev_m,cost_m)], mode="lines+markers",
            line=dict(color=C["blue"],width=2),
            marker=dict(size=5,color=C["blue"],
                        line=dict(color=C["ink"],width=1))))
        mb.update_layout(barmode="group", height=260,
            margin=dict(l=40,r=10,t=10,b=30),
            paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
            font=dict(family=F,size=10,color=C["ink"]),
            legend=dict(orientation="h",y=-0.25),
            xaxis=dict(showgrid=False),
            yaxis=dict(gridcolor=C["ink_l"],tickprefix="$"))

        bud_data = [
            {"Месяц":MONTH_NAMES[i],
             "Выручка $":    f"{rev_m[i]:,.0f}",
             "Бункер $":     f"{bunk_m[i]:,.0f}",
             "Найм $":       f"{hire_m[i]:,.0f}",
             "Риск $":       f"{risk_m[i]:,.0f}",
             "Маржа $":      f"{rev_m[i]-cost_m[i]:,.0f}",
             "Объём mt":     f"{qty_m[i]:,.0f}",
             "Судо-дни":     f"{days_m[i]:.1f}"}
            for i in range(12)
        ]
        bud_data.append({"Месяц":"ИТОГО",
            "Выручка $": f"{sum(rev_m):,.0f}",
            "Бункер $":  f"{sum(bunk_m):,.0f}",
            "Найм $":    f"{sum(hire_m):,.0f}",
            "Риск $":    f"{sum(risk_m):,.0f}",
            "Маржа $":   f"{total_rev-total_cost:,.0f}",
            "Объём mt":  f"{total_qty:,.0f}",
            "Судо-дни":  f"{sum(days_m):.1f}"})

        return html.Div([
            html.Div([
                kpi("Годовая выручка",  money(total_rev)),
                kpi("Годовые затраты",  money(total_cost)),
                kpi("Маржа",            money(total_rev-total_cost), "",
                    C["green"] if total_rev >= total_cost else C["red"]),
                kpi("Объём",            f"{total_qty/1000:,.0f} Kmt",
                    f"{len(shp_res)} отправок"),
                kpi("Ср. затр./тн",     f"${avg_cpt:,.2f}",""),
            ], style={"display":"flex","gap":"8px","flexWrap":"wrap",
                      "marginBottom":"10px"}),
            card([sec("P&L Waterfall"),
                  dcc.Graph(figure=wf, config={"displayModeBar":False})]),
            card([sec("Помесячный план выручка / затраты / маржа"),
                  dcc.Graph(figure=mb, config={"displayModeBar":False})]),
            card([sec("Бюджет по месяцам"),
                  dash_table.DataTable(data=bud_data,
                      columns=[{"name":c,"id":c} for c in bud_data[0].keys()],
                      **TBL)]),
        ])

    # ── СПОТ / РЫНОК ──────────────────────────────────────────────────────
    elif tab == "tab-spot":
        shp_res = (calc or {}).get("shipments", [])
        if not shp_res:
            return card([sec("Спот / Рынок"),
                html.P("Сначала рассчитайте отправки.", style={"color":C["muted"]})])

        rows, own_tce, spot_rates, labels_s, savings = [], [], [], [], []
        for r in shp_res:
            sr = float(next((s.get("spot_rate") or 0
                for s in (shipments or []) if s.get("id")==r.get("shipment_id")),0))
            if sr <= 0: continue
            spot_cost   = sr * r["total_days"] + r["bunker_cost"] + \
                          r["port_cost"] + r["risk_cost"]
            spot_margin = r["revenue"] - spot_cost
            saving      = round(spot_cost - r["total_cost"])
            own_tce.append(r["tce"]); spot_rates.append(sr)
            labels_s.append(f"{r['load'][:10]}→{r['disch'][:10]}")
            savings.append(saving)
            rows.append({
                "ID":               r.get("shipment_id",""),
                "Маршрут":          f"{r['load']} → {r['disch']}",
                "TCE собств. $/d":  f"{r['tce']:,.0f}",
                "Спот $/d":         f"{sr:,.0f}",
                "Затр. (своё) $":   f"{r['total_cost']:,.0f}",
                "Затр. (спот) $":   f"{spot_cost:,.0f}",
                "Маржа (своё) $":   f"{r['margin']:,.0f}",
                "Маржа (спот) $":   f"{spot_margin:,.0f}",
                "Экономия $":       f"{saving:,.0f}",
                "$/тн (своё)":      f"{r['cpt']:,.2f}",
                "Решение":          "Своё" if saving >= 0 else "СПОТ выгоднее",
            })

        if not rows:
            return card([sec("Спот / Рынок"),
                html.P("Укажите «Спот $/d» в колонке отправок и пересчитайте.",
                    style={"color":C["muted"]})])

        # scatter
        mn_ = min(own_tce+spot_rates)*0.9; mx_ = max(own_tce+spot_rates)*1.1
        sc = go.Figure()
        sc.add_shape(type="line",x0=mn_,y0=mn_,x1=mx_,y1=mx_,
            line=dict(color=C["ink_l"],width=1,dash="dash"))
        sc.add_trace(go.Scatter(x=spot_rates, y=own_tce,
            mode="markers+text", text=labels_s, textposition="top center",
            textfont=dict(size=9, color=C["ink"]),
            marker=dict(size=11, color=C["blue_l"],
                line=dict(color=C["ink"],width=1.5))))
        sc.update_layout(height=300,
            margin=dict(l=50,r=10,t=20,b=40),
            paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
            font=dict(family=F,size=10,color=C["ink"]),
                        xaxis=dict(title="Спот $/d",   gridcolor=C["ink_l"]),
            yaxis=dict(title="TCE собств. $/d", gridcolor=C["ink_l"]))

        # savings bar
        colors_s = [C["green_l"] if s >= 0 else C["red_l"] for s in savings]
        sv = go.Figure(go.Bar(x=[r["ID"] for r in rows], y=savings,
            marker_color=colors_s,
            marker_line_color=C["ink"], marker_line_width=1,
            text=[f"${abs(s):,.0f}" for s in savings], textposition="outside"))
        sv.update_layout(height=220, margin=dict(l=40,r=10,t=10,b=40),
            paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
            font=dict(family=F,size=10,color=C["ink"]),
            xaxis=dict(title="Отправка", showgrid=False),
            yaxis=dict(title="Экономия $", gridcolor=C["ink_l"],
                       zeroline=True, zerolinecolor=C["ink"], zerolinewidth=1.5))

        return html.Div([
            card([sec("TCE собственный vs. ставка спот"),
                  html.P("Точки выше диагонали — своё судно дешевле рынка.",
                      style={"fontSize":"11px","color":C["muted"],"marginBottom":"6px"}),
                  dcc.Graph(figure=sc, config={"displayModeBar":False})]),
            card([sec("Экономия (+ своё выгоднее, − спот выгоднее)"),
                  dcc.Graph(figure=sv, config={"displayModeBar":False})]),
            card([dash_table.DataTable(data=rows,
                columns=[{"name":c,"id":c} for c in rows[0].keys()],
                **TBL)]),
        ])

    # ── КАРТА ─────────────────────────────────────────────────────────────
    elif tab == "tab-map":
        shp_res = (calc or {}).get("shipments", [])
        route_layers = []
        all_pts = []
        for r in shp_res:
            pts = []
            for L in r["legs"]:
                if L["from"] in NODES:
                    pts.append([NODES[L["from"]][0], NODES[L["from"]][1]])
            if r["legs"] and r["legs"][-1]["to"] in NODES:
                last = r["legs"][-1]["to"]
                pts.append([NODES[last][0], NODES[last][1]])
            if pts:
                all_pts.extend(pts)
                route_layers.append(dl.Polyline(
                    positions=pts, color=C["blue"], weight=3, opacity=0.8,
                    children=dl.Tooltip(
                        f"{r.get('shipment_id','')} · {r['nm']:,} nm · "
                        f"${r['cpt']:,.2f}/тн")))
            for p in (r["load"], r["disch"]):
                if p in NODES:
                    route_layers.append(dl.CircleMarker(
                        center=[NODES[p][0], NODES[p][1]], radius=6,
                        color=C["ink"], weight=2,
                        fillColor=C["surface"], fillOpacity=1,
                        children=dl.Tooltip(p)))

        legend = html.Div([
            html.Span("  ", style={"display":"inline-block","width":"14px",
                "height":"3px","background":C["blue"],"marginRight":"5px",
                "verticalAlign":"middle"}),
            html.Span("маршруты", style={"fontSize":"11px","color":C["muted"],
                "marginRight":"14px"}),
            html.Span("  ", style={"display":"inline-block","width":"12px",
                "height":"10px","background":C["green_l"],
                "border":f"1px solid {C['green']}","marginRight":"5px",
                "verticalAlign":"middle"}),
            html.Span("SECA", style={"fontSize":"11px","color":C["muted"],
                "marginRight":"14px"}),
            html.Span("  ", style={"display":"inline-block","width":"12px",
                "height":"10px","background":C["amber_l"],
                "border":f"1px solid {C['red']}","marginRight":"5px",
                "verticalAlign":"middle"}),
            html.Span("зоны риска AWRP / LOH / KR",
                style={"fontSize":"11px","color":C["muted"]}),
        ], style={"padding":"8px 0"})

        return html.Div([
            card([sec("Карта маршрутов"),
                  html.P("Показаны все рассчитанные отправки, зоны SECA и зоны риска. "
                         "Цвет маршрута — синий акварельный, контуры чёрные.",
                      style={"fontSize":"11px","color":C["muted"],"marginBottom":"6px"}),
                  legend,
                  dl.Map(center=[25,20], zoom=2,
                      style={"height":"620px","width":"100%",
                             "border":f"1.5px solid {C['ink']}",
                             "borderRadius":"4px"},
                      children=[
                          dl.TileLayer(
                              url="https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}.png",
                              attribution="© OpenStreetMap © CartoDB", maxZoom=12),
                          dl.LayerGroup(seca_rects),
                          dl.LayerGroup(risk_rects),
                          dl.LayerGroup(junction_dots),
                          dl.LayerGroup(route_layers),
                      ])]),
        ])

    # ── ЭКСПОРТ ───────────────────────────────────────────────────────────
    elif tab == "tab-export":
        return html.Div([
            card([sec("Экспорт результатов"),
                  html.P("XLSX содержит листы «Отправки» и «Бюджет». "
                         "PDF — сводный отчёт с таблицей отправок и стоимостью тонны. "
                         "JSON — полный датасет включая справочники.",
                      style={"fontSize":"11px","color":C["muted"],"marginBottom":"10px"}),
                  html.Button("Скачать XLSX", id="btn-xl",
                      style={**btn,"marginBottom":"7px"}),
                  html.Button("Скачать PDF", id="btn-pdf",
                      style={**btn,"marginBottom":"7px","background":C["ink"]}),
                  html.Button("Скачать JSON", id="btn-js",
                      style={**btng,"marginBottom":"0"}),
                  ], style={"maxWidth":"320px"}),
        ])

    return html.Div()


# ── CALLBACKS ─────────────────────────────────────────────────────────────
@app.callback(Output("s-cargoes","data"),
              Input("tbl-cargoes","data"), prevent_initial_call=True)
def sync_cargoes(d): return d or []

@app.callback(Output("s-bunker","data"),
              Input("tbl-bunker","data"), prevent_initial_call=True)
def sync_bunker(d): return d or []

@app.callback(Output("s-vessels","data"),
              Input("tbl-vessels","data"), prevent_initial_call=True)
def sync_vessels(d): return d or []

@app.callback(Output("s-shipments","data"),
              Input("tbl-shipments","data"),
              Input("btn-add-shp","n_clicks"),
              State("s-shipments","data"),
              prevent_initial_call=True)
def sync_shipments(tbl, add_n, store):
    if ctx.triggered_id == "btn-add-shp":
        new = {"id": f"S{len(tbl or [])+1:03d}", "line_id":"L1",
               "cargo_code":"SFO", "load_port":"Pasir Gudang",
               "disch_port":"Rotterdam", "quantity":20000,
               "freight_rate":80, "vessel_id":DEFAULT_VESSELS[0]["name"],
               "contract_type":"Voyage Charter", "spot_rate":14000,
               "month":0, "period_start":"", "period_end":"",
               "density":0.918, "heat_mt_day":0, "status":"planned"}
        return (tbl or []) + [new]
    return tbl or []

@app.callback(Output("s-calc","data"),
              Output("shp-results","children"),
              Input("btn-calc-all","n_clicks"),
              State("s-shipments","data"),
              State("s-vessels","data"),
              State("s-bunker","data"),
              prevent_initial_call=True)
def calc_all(n, shipments, vessels, bunker):
    if not shipments:
        raise PreventUpdate
    vessels = vessels or DEFAULT_VESSELS
    bunker  = bunker  or DEFAULT_BUNKER
    results, errors = [], []
    for shp in shipments:
        vid = shp.get("vessel_id","")
        v   = next((x for x in vessels if x.get("name") == vid), vessels[0])
        r   = calc_shipment(shp, v, bunker)
        if r:
            r["shipment_id"] = shp.get("id","")
            r["line_id"]     = shp.get("line_id","")
            r["vessel"]      = vid
            r["qty"]         = float(shp.get("quantity") or 0)
            r["cargo"]       = shp.get("cargo_code","")
            r["month"]       = int(shp.get("month") or 0)
            results.append(r)
        else:
            errors.append(shp.get("id","?"))

    total_rev  = sum(x["revenue"]    for x in results)
    total_cost = sum(x["total_cost"] for x in results)
    total_qty  = sum(x.get("qty",0)  for x in results)
    avg_cpt    = total_cost / total_qty if total_qty else 0

    summary = html.Div([
        html.Div([
            kpi("Отправок рассчитано", str(len(results))),
            kpi("Суммарная выручка",   money(total_rev)),
            kpi("Суммарные затраты",   money(total_cost)),
            kpi("Маржа",               money(total_rev-total_cost), "",
                C["green"] if total_rev >= total_cost else C["red"]),
            kpi("Ср. затр./тн",        f"${avg_cpt:,.2f}", ""),
        ], style={"display":"flex","gap":"8px","flexWrap":"wrap","marginTop":"10px"}),
        (html.Div(f"Не рассчитано: {', '.join(errors)}",
            style={"color":C["red"],"fontSize":"11px","marginTop":"6px"})
         if errors else html.Div()),
    ])
    return {"shipments": results}, summary


# ── EXPORT CALLBACKS ──────────────────────────────────────────────────────
@app.callback(Output("dl-xlsx","data"),
              Input("btn-xl","n_clicks"),
              State("s-calc","data"),
              State("s-vessels","data"),
              State("s-bunker","data"),
              prevent_initial_call=True)
def dl_xl(n, calc, vessels, bunker):
    data = build_xlsx(calc, vessels or DEFAULT_VESSELS, bunker or DEFAULT_BUNKER)
    return dcc.send_bytes(lambda _: data, "edible_oils_plan.xlsx")

@app.callback(Output("dl-pdf","data"),
              Input("btn-pdf","n_clicks"),
              State("s-calc","data"),
              prevent_initial_call=True)
def dl_pdf(n, calc):
    data = build_pdf(calc)
    return dcc.send_bytes(lambda _: data, "edible_oils_report.pdf")

@app.callback(Output("dl-json","data"),
              Input("btn-js","n_clicks"),
              State("s-calc","data"),
              State("s-shipments","data"),
              State("s-vessels","data"),
              State("s-bunker","data"),
              State("s-cargoes","data"),
              prevent_initial_call=True)
def dl_js(n, calc, shp, vessels, bunker, cargoes):
    payload = {
        "meta": {
            "app": "Edible Oils Charterer's Planner",
            "perspective": "charterer",
            "generated": __import__("datetime").datetime.utcnow().isoformat() + "Z",
        },
        "cargoes":       cargoes or DEFAULT_CARGOES,
        "vessels":       vessels or DEFAULT_VESSELS,
        "bunker_prices": bunker  or DEFAULT_BUNKER,
        "shipments":     shp or [],
        "calc_results":  (calc or {}).get("shipments", []),
        "contract_types": CONTRACT_TYPES,
    }
    data = json.dumps(payload, ensure_ascii=False, indent=2).encode()
    return dcc.send_bytes(lambda _: data, "edible_oils_dataset.json")


# ── LAUNCH ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    app.run(host="0.0.0.0", port=port, debug=False,
            use_reloader=False, dev_tools_hot_reload=False)
