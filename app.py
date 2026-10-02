# ==========================================================================
#  EDIBLE OILS CHARTERER'S PLANNER  (single file)
#  Считаем только затраты фрахтователя и стоимость перевозки тонны, $/т.
#  Направление: грузы — растительные масла наливом.
#  Запуск: pip install dash dash-leaflet openpyxl plotly  → python app.py
# ==========================================================================
import io, json, math, heapq, os, re, base64
from datetime import datetime, timedelta

import dash
from dash import Dash, html, dcc, Input, Output, State, ctx, dash_table, ALL
from dash.exceptions import PreventUpdate
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── PALETTE ───────────────────────────────────────────────────────────────
C = {
    "bg":"#e9edf1","surface":"#ffffff","surf2":"#dfe7ee","ink":"#15202b",
    "ink_l":"#b9c7d3","blue":"#38648a","blue_l":"#c2d8e6","blue_xl":"#e2edf4",
    "muted":"#566877","red":"#8b2635","red_l":"#f3e2e5","amber":"#7a5230",
    "amber_l":"#f3e9da","green":"#2c5739","green_l":"#daece0",
}
F = "-apple-system,'Segoe UI',Roboto,Arial,sans-serif"

def card(ch, **kw):
    return html.Div(ch, style={"background":C["surface"],
        "border":f"1.5px solid {C['ink']}","borderRadius":"4px",
        "padding":"11px 13px","marginBottom":"9px", **kw})

def sec(t, color=None):
    return html.Div(t, style={"fontSize":"10.5px","fontWeight":"800",
        "letterSpacing":"0.07em","textTransform":"uppercase",
        "color":color or C["blue"],"borderBottom":f"1.5px solid {C['ink']}",
        "paddingBottom":"5px","marginBottom":"8px"})

def kpi(title, value, sub="", color=None):
    return html.Div([
        html.Div(title, style={"fontSize":"10px","color":C["muted"],"fontWeight":"700",
            "letterSpacing":"0.05em","textTransform":"uppercase"}),
        html.Div(value, style={"fontSize":"17px","fontWeight":"800",
            "color":color or C["ink"],"marginTop":"2px",
            "fontVariantNumeric":"tabular-nums"}),
        html.Div(sub, style={"fontSize":"10.5px","color":C["muted"],"marginTop":"1px"}),
    ], style={"flex":"1","padding":"9px 11px","background":C["surface"],
        "border":f"1.5px solid {C['ink']}","borderRadius":"4px","minWidth":"130px"})

inp  = {"width":"100%","padding":"5px 7px","boxSizing":"border-box",
        "border":f"1.5px solid {C['ink_l']}","borderRadius":"3px",
        "fontSize":"12px","fontFamily":F,"background":C["surface"]}
btn  = {"padding":"7px 11px","cursor":"pointer","background":C["blue"],
        "color":"#fff","border":f"1.5px solid {C['ink']}","borderRadius":"3px",
        "fontWeight":"700","fontSize":"12px","fontFamily":F,
        "width":"100%","marginTop":"4px"}
btng = {**btn,"background":C["surface"],"color":C["blue"]}
money = lambda x: "$"+f"{float(x):,.0f}"
pmt   = lambda x: "$"+f"{float(x):,.2f}"

TBL = {
    "style_table":{"overflowX":"auto","overflowY":"auto","maxHeight":"340px",
        "border":f"1.5px solid {C['ink']}","borderRadius":"3px"},
    "style_header":{"backgroundColor":C["blue"],"color":"#fff","fontWeight":"700",
        "fontSize":"10.5px","border":f"1px solid {C['ink']}",
        "padding":"4px 6px","position":"sticky","top":"0"},
    "style_cell":{"backgroundColor":C["surface"],"color":C["ink"],"fontSize":"10.5px",
        "padding":"3px 6px","border":f"1px solid {C['ink_l']}","fontFamily":F,
        "minWidth":"62px","maxWidth":"150px","overflow":"hidden",
        "textOverflow":"ellipsis","height":"22px"},
    "style_data_conditional":[{"if":{"row_index":"odd"},"backgroundColor":C["surf2"]}],
}

# ==========================================================================
#  ГЕОГРАФИЯ. Узлы [lat, lon, type]  j = проходная точка, p = порт
# ==========================================================================
NODES = {
    # проходные точки
    "Strait of Gibraltar":[35.950,-5.750,"j"],"Bishop Rock":[49.750,-6.583,"j"],
    "Fastnet Rock":[51.333,-9.600,"j"],"Ile d Ouessant":[48.667,-5.500,"j"],
    "Inishtrahull":[55.417,-7.500,"j"],"Pentland Firth":[58.700,-3.333,"j"],
    "Skagens Odde":[57.800,10.733,"j"],"Nord-Ostsee-Kanal":[54.367,10.150,"j"],
    "Grand Banks South":[42.500,-50.000,"j"],"Montreal":[45.500,-73.550,"j"],
    "Panama":[8.883,-79.517,"j"],"Yucatan Channel":[21.833,-85.050,"j"],
    "Straits of Florida":[24.417,-83.000,"j"],"Punta Arenas":[-53.167,-70.900,"j"],
    "Cape of Good Hope":[-34.367,18.383,"j"],"Cape Leeuwin":[-34.533,115.133,"j"],
    "Selat Sunda":[-6.067,105.833,"j"],"Selat Lombok":[-8.833,115.717,"j"],
    "Selat Wetar":[-8.317,127.450,"j"],"Singapore":[1.267,103.833,"j"],
    "Torres Strait":[-10.550,142.133,"j"],"Wilson Promontory":[-39.167,146.433,"j"],
    "Honshu":[35.000,140.500,"j"],"Tsugaru-Kaikyo":[41.500,140.667,"j"],
    "Port Said":[31.267,32.317,"j"],
    # Европа Балтика / Северное море
    "St. Petersburg":[59.950,30.317,"p"],"Ust-Luga":[59.600,28.400,"p"],
    "Vitino":[67.283,32.350,"p"],"Gdansk":[54.350,18.650,"p"],
    "Hamburg":[53.550,9.967,"p"],"Bremen":[53.083,8.800,"p"],
    "Rotterdam":[51.900,4.467,"p"],"Amsterdam":[52.370,4.900,"p"],
    "Antwerp":[51.217,4.400,"p"],"Ghent":[51.050,3.717,"p"],
    "Dunkirk":[51.050,2.367,"p"],"Le Havre":[49.483,0.117,"p"],
    "London":[51.450,0.367,"p"],"Liverpool":[53.400,-3.000,"p"],
    "Bilbao":[43.350,-3.050,"p"],
    # Атлантика / Иберия / Африка зап.
    "Algeciras":[36.133,-5.433,"p"],"Tanger Med":[35.883,-5.500,"p"],
    "Lisbon":[38.700,-9.133,"p"],"Casablanca":[33.617,-7.617,"p"],
    "Las Palmas":[28.142,-15.417,"p"],"Dakar":[14.683,-17.433,"p"],
    "Abidjan":[5.300,-4.017,"p"],"Tema":[5.633,0.017,"p"],
    "Lagos":[6.433,3.400,"p"],"Douala":[4.050,9.700,"p"],
    "Walvis Bay":[-22.967,14.500,"p"],"Cape Town":[-33.900,18.433,"p"],
    "Saldanha Bay":[-33.017,17.933,"p"],"Durban":[-29.867,31.033,"p"],
    "Richards Bay":[-28.783,32.033,"p"],"Maputo":[-25.967,32.567,"p"],
    "Beira":[-19.833,34.833,"p"],"Dar es Salaam":[-6.817,39.317,"p"],
    "Mombasa":[-4.067,39.667,"p"],"Toamasina":[-18.150,49.417,"p"],
    # Средиземное / Чёрное
    "Valencia":[39.450,-0.333,"p"],"Barcelona":[41.367,2.183,"p"],
    "Marseille":[43.300,5.367,"p"],"Genoa":[44.400,8.933,"p"],
    "Livorno":[43.550,10.300,"p"],"Piraeus":[37.950,23.617,"p"],
    "Thessaloniki":[40.633,22.933,"p"],"Istanbul":[41.017,28.967,"p"],
    "Izmir":[38.417,27.133,"p"],"Mersin":[36.800,34.633,"p"],
    "Iskenderun":[36.583,36.167,"p"],"Constanta":[44.167,28.650,"p"],
    "Burgas":[42.500,27.467,"p"],"Odessa":[46.483,30.733,"p"],
    "Yuzhny":[46.617,31.000,"p"],"Novorossiysk":[44.733,37.783,"p"],
    "Taman":[45.200,36.700,"p"],
    # Левант / Красное море
    "Alexandria":[31.200,29.883,"p"],"Damietta":[31.417,31.817,"p"],
    "Beirut":[33.900,35.517,"p"],"Haifa":[32.833,34.983,"p"],
    "Jeddah":[21.483,39.167,"p"],"Yanbu":[24.083,38.033,"p"],
    "Djibouti":[11.600,43.150,"p"],"Aden":[12.794,44.983,"p"],
    # Залив / Пакистан
    "Fujairah":[25.117,56.333,"p"],"Khor Fakkan":[25.333,56.350,"p"],
    "Jebel Ali":[25.007,55.063,"p"],"Bandar Abbas":[27.183,56.283,"p"],
    "Karachi":[24.800,66.983,"p"],"Port Qasim":[24.783,67.350,"p"],
    # Индия зап.
    "Mumbai":[18.933,72.850,"p"],"Kandla":[23.033,70.217,"p"],
    "Mundra":[22.840,69.717,"p"],"New Mangalore":[12.950,74.800,"p"],
    # Индия вост.
    "Chennai":[13.083,80.267,"p"],"Krishnapatnam":[14.250,80.130,"p"],
    "Kakinada":[16.933,82.233,"p"],"Tuticorin":[8.800,78.150,"p"],
    "Kolkata":[22.567,88.317,"p"],"Haldia":[22.033,88.083,"p"],
    "Chittagong":[22.333,91.833,"p"],"Colombo":[6.950,79.850,"p"],
    # ЮВ Азия
    "Port Klang":[3.000,101.400,"p"],"Pasir Gudang":[1.467,103.900,"p"],
    "Tanjung Pelepas":[1.360,103.550,"p"],"Penang":[5.400,100.317,"p"],
    "Belawan":[3.783,98.683,"p"],"Dumai":[1.667,101.467,"p"],
    "Kuala Tanjung":[2.967,99.450,"p"],"Tanjung Priok":[-6.100,106.883,"p"],
    "Panjang":[-5.467,105.317,"p"],"Surabaya":[-7.200,112.733,"p"],
    "Bangkok":[13.733,100.500,"p"],"Ho Chi Minh":[10.750,106.750,"p"],
    "Manila":[14.600,120.950,"p"],
    # Китай / СВ Азия
    "Hong Kong":[22.300,114.167,"p"],"Nansha":[22.717,113.600,"p"],
    "Xiamen":[24.450,118.067,"p"],"Shanghai":[31.233,121.492,"p"],
    "Ningbo":[29.867,122.133,"p"],"Qingdao":[36.067,120.317,"p"],
    "Tianjin":[38.983,117.733,"p"],"Dalian":[38.917,121.633,"p"],
    "Kaohsiung":[22.617,120.283,"p"],"Busan":[35.100,129.050,"p"],
    "Ulsan":[35.500,129.383,"p"],"Kobe":[34.667,135.200,"p"],
    "Nagoya":[35.050,136.883,"p"],"Yokohama":[35.450,139.650,"p"],
    "Tokyo":[35.633,139.783,"p"],"Vladivostok":[43.108,131.883,"p"],
    # Америка вост.
    "New York":[40.700,-74.017,"p"],"Norfolk":[36.850,-76.300,"p"],
    "Savannah":[32.083,-81.100,"p"],"New Orleans":[29.950,-90.067,"p"],
    "Houston":[29.733,-95.017,"p"],"Corpus Christi":[27.800,-97.400,"p"],
    "Veracruz":[19.200,-96.133,"p"],"Colon":[9.367,-79.917,"p"],
    # Америка зап.
    "Los Angeles":[33.750,-118.271,"p"],"Long Beach":[33.767,-118.200,"p"],
    "Seattle":[47.600,-122.333,"p"],"Vancouver":[49.283,-123.117,"p"],
    "Valparaiso":[-33.033,-71.633,"p"],"Callao":[-12.050,-77.133,"p"],
    # Южная Америка вост.
    "Santos":[-23.950,-46.317,"p"],"Paranagua":[-25.500,-48.517,"p"],
    "Rio Grande":[-32.033,-52.100,"p"],"Montevideo":[-34.900,-56.217,"p"],
    "Buenos Aires":[-34.583,-58.367,"p"],"Rosario":[-32.950,-60.650,"p"],
    # Австралия / Океания
    "Fremantle":[-32.050,115.739,"p"],"Adelaide":[-34.850,138.500,"p"],
    "Melbourne":[-37.817,144.967,"p"],"Sydney":[-33.858,151.217,"p"],
    "Brisbane":[-27.367,153.167,"p"],"Darwin":[-12.467,130.850,"p"],
    "Port Hedland":[-20.317,118.567,"p"],"Honolulu":[21.308,-157.871,"p"],
}

EDGES = [
 ("Strait of Gibraltar","Bishop Rock",969),("Strait of Gibraltar","Fastnet Rock",1068),
 ("Strait of Gibraltar","Grand Banks South",2070),("Strait of Gibraltar","Ile d Ouessant",918),
 ("Strait of Gibraltar","Inishtrahull",1350),("Strait of Gibraltar","Panama",4351),
 ("Strait of Gibraltar","Pentland Firth",1598),("Strait of Gibraltar","Port Said",1943),
 ("Strait of Gibraltar","Cape of Good Hope",5082),("Strait of Gibraltar","Punta Arenas",6352),
 ("Strait of Gibraltar","Straits of Florida",4009),
 ("Bishop Rock","Fastnet Rock",149),("Bishop Rock","Grand Banks South",1830),
 ("Bishop Rock","Ile d Ouessant",80),("Bishop Rock","Inishtrahull",395),
 ("Bishop Rock","Montreal",2844),("Bishop Rock","Nord-Ostsee-Kanal",727),
 ("Bishop Rock","Panama",4388),("Bishop Rock","Pentland Firth",636),
 ("Bishop Rock","Punta Arenas",7019),("Bishop Rock","Skagens Odde",851),
 ("Bishop Rock","Straits of Florida",3859),("Bishop Rock","Yucatan Channel",4176),
 ("Fastnet Rock","Cape of Good Hope",5880),("Fastnet Rock","Grand Banks South",1714),
 ("Fastnet Rock","Ile d Ouessant",225),("Fastnet Rock","Inishtrahull",335),
 ("Fastnet Rock","Panama",4247),("Fastnet Rock","Punta Arenas",7051),
 ("Fastnet Rock","Straits of Florida",3761),("Fastnet Rock","Yucatan Channel",4058),
 ("Grand Banks South","Cape of Good Hope",5954),("Grand Banks South","Ile d Ouessant",1877),
 ("Grand Banks South","Montreal",1335),("Grand Banks South","Panama",2555),
 ("Grand Banks South","Pentland Firth",1998),("Grand Banks South","Straits of Florida",1990),
 ("Ile d Ouessant","Inishtrahull",329),("Ile d Ouessant","Montreal",2901),
 ("Ile d Ouessant","Nord-Ostsee-Kanal",716),("Ile d Ouessant","Panama",4374),
 ("Ile d Ouessant","Pentland Firth",872),("Ile d Ouessant","Punta Arenas",6986),
 ("Ile d Ouessant","Skagens Odde",833),
 ("Inishtrahull","Montreal",2545),("Inishtrahull","Pentland Firth",272),
 ("Inishtrahull","Punta Arenas",7315),
 ("Montreal","Panama",3204),("Montreal","Pentland Firth",2641),
 ("Montreal","Straits of Florida",2531),
 ("Pentland Firth","Nord-Ostsee-Kanal",565),("Pentland Firth","Skagens Odde",450),
 ("Skagens Odde","Nord-Ostsee-Kanal",224),
 ("Port Said","Cape of Good Hope",5340),("Port Said","Cape Leeuwin",6389),
 ("Port Said","Selat Lombok",5892),("Port Said","Selat Sunda",5224),
 ("Port Said","Singapore",5035),("Port Said","Torres Strait",7423),
 ("Cape of Good Hope","Cape Leeuwin",4660),("Cape of Good Hope","Selat Lombok",5461),
 ("Cape of Good Hope","Panama",6466),("Cape of Good Hope","Punta Arenas",4262),
 ("Cape of Good Hope","Singapore",5579),("Cape of Good Hope","Selat Sunda",5164),
 ("Cape of Good Hope","Torres Strait",6854),("Cape of Good Hope","Selat Wetar",6768),
 ("Cape of Good Hope","Wilson Promontory",5530),("Cape of Good Hope","Yucatan Channel",6770),
 ("Cape of Good Hope","Straits of Florida",6770),
 ("Cape Leeuwin","Selat Lombok",1571),("Cape Leeuwin","Selat Sunda",1815),
 ("Cape Leeuwin","Selat Wetar",1954),("Cape Leeuwin","Torres Strait",2717),
 ("Cape Leeuwin","Wilson Promontory",1524),
 ("Selat Lombok","Singapore",963),("Selat Lombok","Selat Sunda",677),
 ("Selat Lombok","Selat Wetar",727),("Selat Lombok","Honshu",3059),
 ("Selat Sunda","Singapore",532),("Selat Sunda","Selat Wetar",1330),("Selat Sunda","Honshu",3171),
 ("Selat Wetar","Singapore",1587),("Selat Wetar","Torres Strait",881),("Selat Wetar","Honshu",2838),
 ("Singapore","Honshu",2879),("Singapore","Tsugaru-Kaikyo",3343),("Singapore","Panama",10505),
 ("Honshu","Panama",7614),("Honshu","Punta Arenas",9286),("Honshu","Torres Strait",3265),
 ("Honshu","Tsugaru-Kaikyo",484),("Honshu","Wilson Promontory",4881),
 ("Tsugaru-Kaikyo","Panama",8004),("Tsugaru-Kaikyo","Punta Arenas",9449),
 ("Torres Strait","Panama",8451),("Torres Strait","Punta Arenas",7217),
 ("Torres Strait","Wilson Promontory",2183),
 ("Wilson Promontory","Panama",7842),("Wilson Promontory","Punta Arenas",5820),
 ("Panama","Yucatan Channel",855),("Panama","Punta Arenas",3932),
 ("Straits of Florida","Yucatan Channel",192),
 # Балтика / Северное море
 ("Rotterdam","Bishop Rock",454),("Rotterdam","Ile d Ouessant",444),
 ("Rotterdam","Nord-Ostsee-Kanal",323),("Rotterdam","Pentland Firth",495),
 ("Rotterdam","Skagens Odde",447),("Rotterdam","Strait of Gibraltar",1371),
 ("Amsterdam","Rotterdam",65),("Antwerp","Rotterdam",55),("Ghent","Antwerp",30),
 ("Le Havre","Bishop Rock",360),("Le Havre","Ile d Ouessant",168),
 ("Dunkirk","Rotterdam",130),("London","Bishop Rock",414),("London","Rotterdam",200),
 ("Liverpool","Fastnet Rock",260),("Hamburg","Nord-Ostsee-Kanal",127),
 ("Hamburg","Skagens Odde",305),("Hamburg","Rotterdam",280),("Bremen","Hamburg",130),
 ("Gdansk","Skagens Odde",530),("Gdansk","Nord-Ostsee-Kanal",400),
 ("St. Petersburg","Skagens Odde",950),("St. Petersburg","Nord-Ostsee-Kanal",800),
 ("Ust-Luga","St. Petersburg",60),("Vitino","St. Petersburg",1380),
 ("Vitino","Skagens Odde",2100),("Vitino","Inishtrahull",1750),
 # Чёрное / Средиземное
 ("Constanta","Istanbul",200),("Odessa","Istanbul",340),("Yuzhny","Odessa",20),
 ("Novorossiysk","Istanbul",500),("Taman","Novorossiysk",50),("Taman","Istanbul",430),
 ("Burgas","Istanbul",160),("Istanbul","Piraeus",340),("Izmir","Piraeus",190),
 ("Thessaloniki","Piraeus",300),("Piraeus","Port Said",590),
 ("Piraeus","Strait of Gibraltar",1050),("Barcelona","Strait of Gibraltar",625),
 ("Barcelona","Marseille",195),("Valencia","Barcelona",180),
 ("Algeciras","Strait of Gibraltar",10),("Tanger Med","Strait of Gibraltar",20),
 ("Genoa","Marseille",185),("Livorno","Genoa",90),("Lisbon","Strait of Gibraltar",320),
 ("Casablanca","Strait of Gibraltar",175),("Casablanca","Las Palmas",515),
 ("Las Palmas","Strait of Gibraltar",681),("Las Palmas","Bishop Rock",1360),
 ("Las Palmas","Straits of Florida",3713),
 # Левант / Красное море
 ("Alexandria","Port Said",180),("Damietta","Port Said",40),("Beirut","Port Said",265),
 ("Haifa","Port Said",260),("Mersin","Port Said",280),("Iskenderun","Mersin",120),
 ("Jeddah","Port Said",660),("Yanbu","Jeddah",200),("Djibouti","Aden",200),
 ("Aden","Port Said",1400),("Aden","Cape of Good Hope",3954),("Aden","Singapore",3631),
 # Залив / Пакистан
 ("Fujairah","Aden",570),("Khor Fakkan","Fujairah",15),("Jebel Ali","Fujairah",70),
 ("Bandar Abbas","Fujairah",130),("Karachi","Fujairah",660),("Port Qasim","Karachi",25),
 # Индия зап.
 ("Mumbai","Karachi",532),("Kandla","Mumbai",390),("Mundra","Kandla",60),
 ("New Mangalore","Mumbai",340),("New Mangalore","Colombo",450),
 ("New Mangalore","Kandla",700),
 # Индия вост.
 ("Chennai","Mumbai",790),("Chennai","Krishnapatnam",170),
 ("Krishnapatnam","Kakinada",230),("Krishnapatnam","Colombo",500),
 ("Kakinada","Chennai",360),("Kakinada","Kolkata",540),
 ("Tuticorin","Chennai",270),("Tuticorin","Colombo",140),
 ("Kolkata","Chennai",870),("Haldia","Kolkata",60),("Chittagong","Kolkata",390),
 ("Colombo","Chennai",175),("Colombo","Singapore",1581),
 ("Colombo","Port Said",3481),("Colombo","Cape of Good Hope",4317),
 # ЮВ Азия
 ("Port Klang","Singapore",200),("Pasir Gudang","Singapore",30),
 ("Tanjung Pelepas","Singapore",40),("Penang","Port Klang",230),
 ("Belawan","Singapore",240),("Dumai","Singapore",150),("Kuala Tanjung","Dumai",80),
 ("Tanjung Priok","Singapore",560),("Panjang","Tanjung Priok",150),
 ("Surabaya","Tanjung Priok",420),("Bangkok","Singapore",835),
 ("Ho Chi Minh","Singapore",650),("Manila","Hong Kong",640),
 # Китай / СВ Азия
 ("Hong Kong","Singapore",1460),("Nansha","Hong Kong",110),("Xiamen","Hong Kong",300),
 ("Shanghai","Hong Kong",770),("Ningbo","Shanghai",130),("Qingdao","Shanghai",390),
 ("Tianjin","Shanghai",790),("Dalian","Tianjin",250),("Kaohsiung","Hong Kong",340),
 ("Busan","Hong Kong",780),("Busan","Honshu",450),("Ulsan","Busan",50),
 ("Kobe","Busan",380),("Kobe","Honshu",50),("Nagoya","Kobe",130),("Nagoya","Honshu",100),
 ("Yokohama","Honshu",20),("Yokohama","Nagoya",200),("Tokyo","Yokohama",15),
 ("Shanghai","Honshu",550),("Vladivostok","Honshu",620),
 # Америка
 ("New York","Strait of Gibraltar",3180),("New York","Panama",2016),
 ("New York","Grand Banks South",950),("New York","Straits of Florida",1087),
 ("Norfolk","New York",280),("Savannah","Norfolk",560),
 ("New Orleans","Straits of Florida",900),("Houston","New Orleans",300),
 ("Houston","Straits of Florida",1200),("Houston","Panama",1150),
 ("Corpus Christi","Houston",200),("Veracruz","Houston",750),("Colon","Panama",44),
 ("Los Angeles","Panama",2913),("Long Beach","Los Angeles",15),
 ("Seattle","Vancouver",85),("Vancouver","Los Angeles",979),
 ("Valparaiso","Punta Arenas",1500),("Valparaiso","Panama",2983),
 ("Callao","Valparaiso",1274),("Callao","Panama",1310),
 # Юж. Америка вост.
 ("Santos","Strait of Gibraltar",4690),("Santos","Cape of Good Hope",3560),
 ("Paranagua","Santos",180),("Rio Grande","Santos",620),
 ("Buenos Aires","Santos",1080),("Montevideo","Buenos Aires",120),
 ("Rosario","Buenos Aires",220),("Buenos Aires","Punta Arenas",1394),
 ("Buenos Aires","Cape of Good Hope",3704),
 # Африка
 ("Dakar","Las Palmas",930),("Dakar","Strait of Gibraltar",1460),
 ("Abidjan","Dakar",950),("Tema","Abidjan",220),("Lagos","Tema",250),
 ("Douala","Lagos",350),("Walvis Bay","Cape Town",700),
 ("Cape Town","Cape of Good Hope",35),("Saldanha Bay","Cape Town",100),
 ("Durban","Cape of Good Hope",764),("Richards Bay","Durban",90),
 ("Maputo","Durban",220),("Beira","Maputo",500),
 ("Mombasa","Durban",1390),("Mombasa","Aden",1450),
 ("Dar es Salaam","Mombasa",220),("Toamasina","Durban",1300),
 # Австралия
 ("Fremantle","Cape Leeuwin",181),("Fremantle","Cape of Good Hope",4755),
 ("Fremantle","Selat Lombok",1440),("Port Hedland","Fremantle",870),
 ("Darwin","Singapore",1840),("Adelaide","Fremantle",1130),
 ("Melbourne","Adelaide",430),("Sydney","Melbourne",560),
 ("Sydney","Torres Strait",1670),("Sydney","Wilson Promontory",440),
 ("Brisbane","Sydney",440),
]

ADJ = {}
for _a, _b, _w in EDGES:
    ADJ.setdefault(_a, []).append((_b, _w))
    ADJ.setdefault(_b, []).append((_a, _w))

# ── ЗОНЫ ──────────────────────────────────────────────────────────────────
SEASONAL = [
 {"n":"North Atlantic Winter","m":[11,12,1,2,3],"a":40,"b":65,"c":-65,"d":20,"f":0.87},
 {"n":"North Sea Winter","m":[10,11,12,1,2,3],"a":48,"b":65,"c":-5,"d":15,"f":0.88},
 {"n":"Baltic Winter","m":[12,1,2,3],"a":53,"b":66,"c":9,"d":32,"f":0.85},
 {"n":"Arctic White Sea Winter","m":[10,11,12,1,2,3,4],"a":62,"b":75,"c":5,"d":45,"f":0.78},
 {"n":"Bengal Monsoon","m":[6,7,8,9],"a":5,"b":22,"c":78,"d":98,"f":0.82},
 {"n":"Arabian Sea Monsoon","m":[6,7,8],"a":5,"b":25,"c":50,"d":78,"f":0.80},
 {"n":"South Indian Winter","m":[6,7,8,9],"a":-60,"b":-35,"c":20,"d":115,"f":0.83},
 {"n":"NW Pacific Typhoons","m":[7,8,9,10],"a":10,"b":35,"c":110,"d":165,"f":0.85},
 {"n":"North Pacific Winter","m":[11,12,1,2,3],"a":35,"b":60,"c":145,"d":180,"f":0.85},
 {"n":"Cape Storms","m":[5,6,7,8,9],"a":-40,"b":-28,"c":10,"d":30,"f":0.84},
 {"n":"Red Sea Summer","m":[5,6,7,8,9],"a":12,"b":30,"c":32,"d":44,"f":0.93},
 {"n":"South China Sea NE Monsoon","m":[11,12,1,2],"a":5,"b":25,"c":105,"d":125,"f":0.88},
]
SECA = [
 {"n":"Baltic ECA","a":53.5,"b":66.0,"c":9.0,"d":30.0},
 {"n":"North Sea ECA","a":48.0,"b":62.0,"c":-5.0,"d":9.0},
 {"n":"North American ECA","a":24.0,"b":50.0,"c":-140.0,"d":-50.0},
 {"n":"US Caribbean ECA","a":15.0,"b":31.0,"c":-98.0,"d":-60.0},
 {"n":"Chinese ECA","a":18.0,"b":41.0,"c":107.0,"d":125.0},
 {"n":"Mediterranean ECA","a":30.0,"b":46.0,"c":-6.0,"d":36.0},
]
RISK = [
 {"n":"Gulf of Aden / Somali Basin","t":"AWRP+LOH","a":-5,"b":16,"c":45,"d":65,"aw":0.035,"lo":0.025},
 {"n":"Red Sea","t":"AWRP+LOH","a":12,"b":30,"c":32,"d":44,"aw":0.050,"lo":0.035},
 {"n":"Persian Gulf","t":"AWRP+LOH","a":22,"b":30,"c":48,"d":57,"aw":0.020,"lo":0.015},
 {"n":"Black Sea","t":"AWRP+LOH","a":41,"b":47.5,"c":28,"d":37.5,"aw":0.120,"lo":0.080},
 {"n":"Korean Waters","t":"KR","a":32,"b":42,"c":124,"d":132,"kr":3000},
 {"n":"Taiwan Strait","t":"KR","a":22,"b":27,"c":118,"d":122.5,"kr":1500},
]

# ==========================================================================
#  ДВИЖОК
# ==========================================================================
def haversine(a, b):
    R = 3440.065
    p1, p2 = math.radians(a[0]), math.radians(b[0])
    dp = p2-p1; dl = math.radians(b[1]-a[1])
    h = math.sin(dp/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 2*R*math.asin(min(1.0, math.sqrt(h)))

def dijkstra(s, d):
    if s not in NODES or d not in NODES: return None
    if s == d: return {"total":0.0, "legs":[]}
    dist = {s:0.0}; prev = {}; h = [(0.0,s)]; done = set()
    while h:
        dd, u = heapq.heappop(h)
        if u in done: continue
        done.add(u)
        if u == d: break
        for v, w in ADJ.get(u, []):
            nd = dd+w
            if nd < dist.get(v, math.inf):
                dist[v] = nd; prev[v] = (u,w); heapq.heappush(h,(nd,v))
    if d not in dist: return None
    legs = []; cur = d
    while cur != s:
        p, w = prev[cur]; legs.append({"from":p,"to":cur,"nm":w}); cur = p
    legs.reverse()
    return {"total":dist[d], "legs":legs}

def in_box(lat, lon, z):
    return z["a"] <= lat <= z["b"] and z["c"] <= lon <= z["d"]

def hits(na, nb, z):
    mid = ((na[0]+nb[0])/2, (na[1]+nb[1])/2)
    return any(in_box(p[0],p[1],z) for p in (na,nb,mid))

def season_factor(na, nb, month):
    if not month: return 1.0, []
    fs, ns = [], []
    for r in SEASONAL:
        if month in r["m"] and hits(na, nb, r):
            fs.append(r["f"]); ns.append(r["n"])
    return (min(fs) if fs else 1.0), ns

def seca_names(na, nb):  return [z["n"] for z in SECA if hits(na,nb,z)]
def risk_zones(na, nb):  return [z for z in RISK if hits(na,nb,z)]

def bunker_price(port, fuel, ref):
    for b in ref:
        if b["port"] == port:
            return float(b.get(fuel,0) or 0)
    return {"vlsfo":590,"mgo":930}.get(fuel,590)

def port_cost(port, ref):
    for b in ref:
        if b["port"] == port:
            return (float(b.get("port_dues",0) or 0),
                    float(b.get("canal_dues",0) or 0),
                    float(b.get("agents",0) or 0))
    return 18000, 0, 7000

def voyage_cost(load, disch, qty, month, vessel, bunker_ref,
                contract="TC In", canal_override=None):
    """
    Затраты фрахтователя по одному переходу. Только затраты, без маржи.
    Возвращает словарь с компонентами и стоимостью тонны.
    """
    if load not in NODES or disch not in NODES: return None
    r = dijkstra(load, disch)
    if not r or not r["legs"]: return None

    spd     = float(vessel.get("speed") or 13)
    cons_s  = float(vessel.get("cons_sea") or 26)
    cons_m  = float(vessel.get("cons_mgo") or 19)
    cons_p  = float(vessel.get("cons_port") or 3)
    tank    = float(vessel.get("tank") or 1100)
    load_h  = float(vessel.get("load_hours") or 30)
    disch_h = float(vessel.get("disch_hours") or 36)
    value   = float(vessel.get("value") or 22_000_000)
    tc_rate = float(vessel.get("tc_rate") or 14500)
    opex    = float(vessel.get("opex_day") or 7000)
    heat_mt = float(vessel.get("heat_mt_day") or 0)

    rob = tank
    ifo = mgo = days_sea = nm = 0.0
    awrp = loh = kr = 0.0
    seca_set, risk_set = set(), []
    alerts = []
    sfmin = 1.0

    for L in r["legs"]:
        na, nb = NODES[L["from"]], NODES[L["to"]]
        sf, _ = season_factor(na, nb, month)
        sfmin = min(sfmin, sf)
        eff = spd*sf
        dz = L["nm"]/(eff*24) if eff > 0 else 0
        sz = seca_names(na, nb)
        if sz:
            need = cons_m*dz
            if rob < need*1.05:
                alerts.append({"port":L["from"],"mt":round(need*1.05-rob,1),"zone":sz[0]})
            mgo += need; rob = max(0.0, rob-need); seca_set.update(sz)
        else:
            need = cons_s*dz; ifo += need; rob = max(0.0, rob-need)
        days_sea += dz; nm += L["nm"]
        for rz in risk_zones(na, nb):
            risk_set.append(rz)
            awrp += rz.get("aw",0)*value/100
            loh  += rz.get("lo",0)*value/100
            kr   += rz.get("kr",0)*dz

    port_days = (load_h+disch_h)/24
    ifo += cons_p*port_days
    total_days = days_sea+port_days

    vlsfo_p = bunker_price(load,"vlsfo",bunker_ref)
    mgo_p   = bunker_price(load,"mgo",bunker_ref)
    heat_cost = heat_mt*total_days*qty/1000*vlsfo_p
    bunker = ifo*vlsfo_p + mgo*mgo_p + heat_cost

    pd_l, cd_l, ag_l = port_cost(load, bunker_ref)
    pd_d, cd_d, ag_d = port_cost(disch, bunker_ref)
    ports = pd_l+pd_d+ag_l+ag_d+cd_l+cd_d
    if canal_override:
        ports += float(canal_override)

    risk = awrp+loh+kr

    if contract in ("TC In","TC Out","Spot TC"):
        hire = tc_rate*total_days; opex_c = 0.0
    elif contract == "Own":
        hire = 0.0; opex_c = opex*total_days
    else:                        # Voyage Charter, CoA, Lastage
        hire = 0.0; opex_c = opex*total_days*0.5

    total = bunker+ports+risk+hire+opex_c
    cpt   = total/qty if qty else 0

    return {
        "load":load,"disch":disch,"nm":round(nm),
        "days_sea":round(days_sea,2),"port_days":round(port_days,2),
        "total_days":round(total_days,2),
        "ifo":round(ifo,1),"mgo":round(mgo,1),"heat_cost":round(heat_cost),
        "bunker":round(bunker),"ports":round(ports),
        "hire":round(hire),"opex":round(opex_c),"risk":round(risk),
        "total":round(total),"cpt":round(cpt,2),
        "sfactor":round(sfmin,3),"seca":sorted(seca_set),
        "risk_zones":[z["n"] for z in risk_set],
        "alerts":alerts,"vlsfo_price":vlsfo_p,"mgo_price":mgo_p,
        "legs":r["legs"],
    }

# ==========================================================================
#  СПРАВОЧНИКИ ПО УМОЛЧАНИЮ
# ==========================================================================
DEFAULT_CARGOES = [
 {"code":"SFO","name":"Подсолнечное масло","group":"Soft oil","density":0.918,
  "heat_deg":0,"heat_mt":0,"parcel_min":1000,"parcel_max":35000,
  "tanks":"coated","notes":"FFA sensitive"},
 {"code":"SBO","name":"Соевое масло","group":"Soft oil","density":0.920,
  "heat_deg":0,"heat_mt":0,"parcel_min":1000,"parcel_max":40000,
  "tanks":"coated","notes":"Parcel tanker grade"},
 {"code":"RSO","name":"Рапсовое масло","group":"Soft oil","density":0.914,
  "heat_deg":0,"heat_mt":0,"parcel_min":1000,"parcel_max":30000,
  "tanks":"coated","notes":"EU biodiesel feedstock"},
 {"code":"CPO","name":"Пальмовое масло сырое","group":"Palm","density":0.891,
  "heat_deg":50,"heat_mt":1.8,"parcel_min":3000,"parcel_max":50000,
  "tanks":"stainless or coated","notes":"50-55 C heating"},
 {"code":"RBDPO","name":"Пальмовое масло RBD","group":"Palm","density":0.891,
  "heat_deg":45,"heat_mt":1.5,"parcel_min":2000,"parcel_max":50000,
  "tanks":"stainless","notes":"Edible grade"},
 {"code":"POL","name":"Пальмовый олеин","group":"Palm fraction","density":0.910,
  "heat_deg":40,"heat_mt":1.2,"parcel_min":2000,"parcel_max":45000,
  "tanks":"stainless","notes":"Cloud point controlled"},
 {"code":"PST","name":"Пальмовый стеарин","group":"Palm fraction","density":0.880,
  "heat_deg":60,"heat_mt":2.4,"parcel_min":2000,"parcel_max":40000,
  "tanks":"stainless","notes":"Keep 60-65 C"},
 {"code":"PFAD","name":"Пальмовые жирные кислоты","group":"Palm fraction",
  "density":0.870,"heat_deg":60,"heat_mt":2.6,"parcel_min":2000,"parcel_max":40000,
  "tanks":"stainless","notes":"Corrosive"},
 {"code":"PKO","name":"Пальмоядровое масло","group":"Lauric","density":0.905,
  "heat_deg":35,"heat_mt":1.0,"parcel_min":1000,"parcel_max":25000,
  "tanks":"stainless","notes":"Lauric grade"},
 {"code":"CNO","name":"Кокосовое масло","group":"Lauric","density":0.908,
  "heat_deg":35,"heat_mt":1.0,"parcel_min":1000,"parcel_max":20000,
  "tanks":"stainless","notes":"Lauric grade"},
]

DEFAULT_VESSELS = [
 {"name":"MT Elaeis Trader","cls":"IMO 2 parcel 20k","intake":19900,"speed":14.0,
  "cons_sea":26,"cons_mgo":19,"cons_port":3.0,"tank":1100,"load_hours":30,
  "disch_hours":38,"status":"TC In","tc_rate":9200,"opex_day":7000,
  "value":22000000,"offhire_days":12,"heat_mt_day":2.0,"notes":"28 stainless tanks"},
 {"name":"MT Olea Star","cls":"IMO 2 parcel 13k","intake":12800,"speed":13.0,
  "cons_sea":20,"cons_mgo":15,"cons_port":2.4,"tank":760,"load_hours":26,
  "disch_hours":34,"status":"Own","tc_rate":8400,"opex_day":6200,
  "value":16000000,"offhire_days":14,"heat_mt_day":1.6,"notes":"14 coated tanks"},
 {"name":"MT Palm Express","cls":"IMO 2 parcel 25k","intake":24500,"speed":14.5,
  "cons_sea":30,"cons_mgo":22,"cons_port":3.5,"tank":1400,"load_hours":36,
  "disch_hours":44,"status":"TC In","tc_rate":16800,"opex_day":0,
  "value":32000000,"offhire_days":8,"heat_mt_day":2.4,"notes":"22 stainless tanks"},
 {"name":"MT Solaris","cls":"IMO 2 parcel 37k","intake":36800,"speed":14.0,
  "cons_sea":36,"cons_mgo":26,"cons_port":4.0,"tank":1800,"load_hours":48,
  "disch_hours":56,"status":"Own","tc_rate":11200,"opex_day":8500,
  "value":44000000,"offhire_days":10,"heat_mt_day":3.2,"notes":"33 tanks stainless"},
 {"name":"MT Oleander","cls":"IMO 2 chemical 8k","intake":7900,"speed":12.5,
  "cons_sea":16,"cons_mgo":12,"cons_port":2.0,"tank":560,"load_hours":22,
  "disch_hours":28,"status":"Own","tc_rate":7800,"opex_day":5800,
  "value":12000000,"offhire_days":18,"heat_mt_day":1.2,"notes":"18 tanks"},
]

DEFAULT_BUNKER = [
 {"port":"Rotterdam","vlsfo":585,"mgo":905,"port_dues":44000,"canal_dues":0,"agents":9000},
 {"port":"Amsterdam","vlsfo":588,"mgo":912,"port_dues":41000,"canal_dues":0,"agents":8500},
 {"port":"Antwerp","vlsfo":592,"mgo":918,"port_dues":40000,"canal_dues":0,"agents":8800},
 {"port":"Hamburg","vlsfo":598,"mgo":925,"port_dues":38000,"canal_dues":0,"agents":8200},
 {"port":"Le Havre","vlsfo":600,"mgo":930,"port_dues":39000,"canal_dues":0,"agents":8000},
 {"port":"Ghent","vlsfo":594,"mgo":920,"port_dues":36000,"canal_dues":0,"agents":8200},
 {"port":"St. Petersburg","vlsfo":620,"mgo":965,"port_dues":30000,"canal_dues":0,"agents":11000},
 {"port":"Ust-Luga","vlsfo":622,"mgo":967,"port_dues":28000,"canal_dues":0,"agents":10500},
 {"port":"Vitino","vlsfo":640,"mgo":990,"port_dues":22000,"canal_dues":0,"agents":14000},
 {"port":"Gdansk","vlsfo":610,"mgo":955,"port_dues":28000,"canal_dues":0,"agents":9500},
 {"port":"Novorossiysk","vlsfo":610,"mgo":960,"port_dues":26000,"canal_dues":0,"agents":11000},
 {"port":"Taman","vlsfo":612,"mgo":962,"port_dues":24000,"canal_dues":0,"agents":10500},
 {"port":"Odessa","vlsfo":615,"mgo":965,"port_dues":27000,"canal_dues":0,"agents":12000},
 {"port":"Constanta","vlsfo":608,"mgo":955,"port_dues":25000,"canal_dues":0,"agents":9500},
 {"port":"Istanbul","vlsfo":605,"mgo":950,"port_dues":22000,"canal_dues":0,"agents":9000},
 {"port":"Piraeus","vlsfo":602,"mgo":945,"port_dues":24000,"canal_dues":0,"agents":8800},
 {"port":"Port Said","vlsfo":598,"mgo":940,"port_dues":18000,"canal_dues":180000,"agents":12000},
 {"port":"Alexandria","vlsfo":600,"mgo":940,"port_dues":21000,"canal_dues":0,"agents":8600},
 {"port":"Damietta","vlsfo":602,"mgo":942,"port_dues":20000,"canal_dues":0,"agents":8400},
 {"port":"Jeddah","vlsfo":612,"mgo":965,"port_dues":32000,"canal_dues":0,"agents":10000},
 {"port":"Fujairah","vlsfo":590,"mgo":935,"port_dues":30000,"canal_dues":0,"agents":9500},
 {"port":"Jebel Ali","vlsfo":592,"mgo":937,"port_dues":31000,"canal_dues":0,"agents":9500},
 {"port":"Karachi","vlsfo":618,"mgo":975,"port_dues":24000,"canal_dues":0,"agents":9000},
 {"port":"Port Qasim","vlsfo":620,"mgo":977,"port_dues":22000,"canal_dues":0,"agents":8800},
 {"port":"Mumbai","vlsfo":622,"mgo":980,"port_dues":28000,"canal_dues":0,"agents":9500},
 {"port":"Kandla","vlsfo":625,"mgo":985,"port_dues":23000,"canal_dues":0,"agents":8800},
 {"port":"Mundra","vlsfo":624,"mgo":984,"port_dues":22000,"canal_dues":0,"agents":8600},
 {"port":"New Mangalore","vlsfo":626,"mgo":986,"port_dues":20000,"canal_dues":0,"agents":8400},
 {"port":"Chennai","vlsfo":620,"mgo":978,"port_dues":25000,"canal_dues":0,"agents":9000},
 {"port":"Krishnapatnam","vlsfo":621,"mgo":979,"port_dues":19000,"canal_dues":0,"agents":8200},
 {"port":"Kakinada","vlsfo":622,"mgo":980,"port_dues":18000,"canal_dues":0,"agents":8000},
 {"port":"Tuticorin","vlsfo":618,"mgo":976,"port_dues":17000,"canal_dues":0,"agents":7800},
 {"port":"Kolkata","vlsfo":622,"mgo":980,"port_dues":26000,"canal_dues":0,"agents":9200},
 {"port":"Haldia","vlsfo":624,"mgo":982,"port_dues":24000,"canal_dues":0,"agents":8800},
 {"port":"Chittagong","vlsfo":628,"mgo":988,"port_dues":27000,"canal_dues":0,"agents":9800},
 {"port":"Colombo","vlsfo":610,"mgo":960,"port_dues":20000,"canal_dues":0,"agents":8000},
 {"port":"Singapore","vlsfo":578,"mgo":915,"port_dues":34000,"canal_dues":0,"agents":9000},
 {"port":"Port Klang","vlsfo":582,"mgo":920,"port_dues":28000,"canal_dues":0,"agents":8200},
 {"port":"Pasir Gudang","vlsfo":580,"mgo":918,"port_dues":27000,"canal_dues":0,"agents":8000},
 {"port":"Belawan","vlsfo":592,"mgo":930,"port_dues":24000,"canal_dues":0,"agents":7600},
 {"port":"Dumai","vlsfo":590,"mgo":928,"port_dues":22000,"canal_dues":0,"agents":7400},
 {"port":"Kuala Tanjung","vlsfo":591,"mgo":929,"port_dues":21000,"canal_dues":0,"agents":7200},
 {"port":"Tanjung Priok","vlsfo":588,"mgo":925,"port_dues":25000,"canal_dues":0,"agents":7800},
 {"port":"Bangkok","vlsfo":596,"mgo":934,"port_dues":26000,"canal_dues":0,"agents":8200},
 {"port":"Ho Chi Minh","vlsfo":594,"mgo":932,"port_dues":24000,"canal_dues":0,"agents":8000},
 {"port":"Hong Kong","vlsfo":600,"mgo":940,"port_dues":36000,"canal_dues":0,"agents":9500},
 {"port":"Shanghai","vlsfo":602,"mgo":942,"port_dues":35000,"canal_dues":0,"agents":9200},
 {"port":"Nansha","vlsfo":600,"mgo":940,"port_dues":33000,"canal_dues":0,"agents":9000},
 {"port":"Tianjin","vlsfo":608,"mgo":948,"port_dues":36000,"canal_dues":0,"agents":9500},
 {"port":"Qingdao","vlsfo":606,"mgo":946,"port_dues":34000,"canal_dues":0,"agents":9200},
 {"port":"Busan","vlsfo":596,"mgo":936,"port_dues":34000,"canal_dues":0,"agents":9000},
 {"port":"Kobe","vlsfo":610,"mgo":950,"port_dues":38000,"canal_dues":0,"agents":10000},
 {"port":"Yokohama","vlsfo":612,"mgo":952,"port_dues":39000,"canal_dues":0,"agents":10200},
 {"port":"New York","vlsfo":605,"mgo":945,"port_dues":48000,"canal_dues":0,"agents":12000},
 {"port":"Houston","vlsfo":595,"mgo":935,"port_dues":42000,"canal_dues":0,"agents":11000},
 {"port":"New Orleans","vlsfo":592,"mgo":932,"port_dues":40000,"canal_dues":0,"agents":10500},
 {"port":"Los Angeles","vlsfo":608,"mgo":948,"port_dues":50000,"canal_dues":0,"agents":12500},
 {"port":"Santos","vlsfo":618,"mgo":972,"port_dues":32000,"canal_dues":0,"agents":11000},
 {"port":"Paranagua","vlsfo":620,"mgo":974,"port_dues":30000,"canal_dues":0,"agents":10500},
 {"port":"Rosario","vlsfo":622,"mgo":978,"port_dues":28000,"canal_dues":0,"agents":12000},
 {"port":"Buenos Aires","vlsfo":620,"mgo":975,"port_dues":30000,"canal_dues":0,"agents":11500},
 {"port":"Las Palmas","vlsfo":588,"mgo":920,"port_dues":22000,"canal_dues":0,"agents":8000},
 {"port":"Dakar","vlsfo":606,"mgo":950,"port_dues":20000,"canal_dues":0,"agents":8500},
 {"port":"Mombasa","vlsfo":616,"mgo":968,"port_dues":24000,"canal_dues":0,"agents":9000},
 {"port":"Durban","vlsfo":612,"mgo":962,"port_dues":26000,"canal_dues":0,"agents":9200},
 {"port":"Cape Town","vlsfo":610,"mgo":960,"port_dues":25000,"canal_dues":0,"agents":9000},
 {"port":"Fremantle","vlsfo":610,"mgo":960,"port_dues":22000,"canal_dues":0,"agents":9000},
 {"port":"Sydney","vlsfo":618,"mgo":972,"port_dues":28000,"canal_dues":0,"agents":9500},
 {"port":"Panama","vlsfo":598,"mgo":938,"port_dues":30000,"canal_dues":240000,"agents":14000},
]

CONTRACT_TYPES = ["TC In","TC Out","Spot TC","Own","Voyage Charter","CoA","Lastage"]
MONTH_NAMES    = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

# стандартные маршруты-шаблоны
DEFAULT_ROUTES = [
 {"id":"R01","name":"Индонезия - Роттердам",
  "load":"Pasir Gudang","disch":"Rotterdam","cargo":"CPO","notes":"CPO/PKO/CNO"},
 {"id":"R02","name":"Малайзия - Роттердам",
  "load":"Kuala Tanjung","disch":"Rotterdam","cargo":"RBDPO","notes":"Palm products"},
 {"id":"R03","name":"Аргентина - Роттердам",
  "load":"Rosario","disch":"Rotterdam","cargo":"SBO","notes":"SBO/SFO season"},
 {"id":"R04","name":"Украина - Индия",
  "load":"Yuzhny","disch":"Mumbai","cargo":"SFO","notes":"SFO/RSO Black Sea"},
 {"id":"R05","name":"Россия ЧМ - Индия",
  "load":"Novorossiysk","disch":"Mumbai","cargo":"SFO","notes":"SFO Black Sea"},
 {"id":"R06","name":"Индонезия - Индия зап.",
  "load":"Tanjung Priok","disch":"Kandla","cargo":"CPO","notes":"CPO/PKO"},
 {"id":"R07","name":"Индонезия - Китай",
  "load":"Pasir Gudang","disch":"Tianjin","cargo":"RBDPO","notes":"Lauric/Palm"},
 {"id":"R08","name":"Аргентина - Индия",
  "load":"Rosario","disch":"Kandla","cargo":"SBO","notes":"SBO via Cape"},
 {"id":"R09","name":"Малайзия - Стамбул",
  "load":"Pasir Gudang","disch":"Istanbul","cargo":"CPO","notes":"Palm Med"},
 {"id":"R10","name":"Россия - Антверп",
  "load":"Novorossiysk","disch":"Antwerp","cargo":"SFO","notes":"SFO/RSO"},
]

# ==========================================================================
#  КАРТА — статичная Scattergeo через Plotly
# ==========================================================================
REGION_BOUNDS = {
    "Мир целиком":       {"lat":[0],"lon":[20],"zoom":1},
    "Европа / ЧМ":       {"lat":[50],"lon":[20],"zoom":3},
    "ЮВ Азия":           {"lat":[5],"lon":[110],"zoom":3},
    "Индийский океан":   {"lat":[10],"lon":[70],"zoom":2},
    "Атлантика":         {"lat":[20],"lon":[-30],"zoom":2},
    "Ближний Восток":    {"lat":[20],"lon":[50],"zoom":3},
    "Китай / Япония":    {"lat":[35],"lon":[120],"zoom":3},
    "Южная Америка":     {"lat":[-25],"lon":[-55],"zoom":2},
    "Африка":            {"lat":[-5],"lon":[25],"zoom":2},
    "Австралия":         {"lat":[-30],"lon":[135],"zoom":3},
}

def build_map(shipment_results=None, region="Мир целиком"):
    fig = go.Figure()
    rb = REGION_BOUNDS.get(region, REGION_BOUNDS["Мир целиком"])

    # зоны SECA — прямоугольники через линии
    for z in SECA:
        lats = [z["a"],z["a"],z["b"],z["b"],z["a"],None]
        lons = [z["c"],z["d"],z["d"],z["c"],z["c"],None]
        fig.add_trace(go.Scattergeo(
            lat=lats, lon=lons, mode="lines",
            line=dict(color="#2c5739", width=1.5, dash="dot"),
            fill="toself", fillcolor="rgba(44,87,57,0.08)",
            name="SECA", legendgroup="seca",
            showlegend=z == SECA[0],
            hovertemplate=f"<b>SECA</b><br>{z['n']}<extra></extra>",
        ))

    # зоны риска
    for z in RISK:
        lats = [z["a"],z["a"],z["b"],z["b"],z["a"],None]
        lons = [z["c"],z["d"],z["d"],z["c"],z["c"],None]
        fig.add_trace(go.Scattergeo(
            lat=lats, lon=lons, mode="lines",
            line=dict(color="#8b2635", width=1.5, dash="dash"),
            fill="toself", fillcolor="rgba(139,38,53,0.09)",
            name="Risk Zone", legendgroup="risk",
            showlegend=z == RISK[0],
            hovertemplate=f"<b>Risk: {z['t']}</b><br>{z['n']}<extra></extra>",
        ))

    # маршруты из результатов
    if shipment_results:
        for r in shipment_results:
            pts = []
            for L in r.get("legs", []):
                if L["from"] in NODES:
                    pts.append((NODES[L["from"]][0], NODES[L["from"]][1]))
            if r["legs"] and r["legs"][-1]["to"] in NODES:
                n = r["legs"][-1]["to"]
                pts.append((NODES[n][0], NODES[n][1]))
            if len(pts) < 2: continue
            lats = [p[0] for p in pts]
            lons = [p[1] for p in pts]
            label = (f"{r.get('shipment_id','')} · "
                     f"{r['load']} → {r['disch']}<br>"
                     f"{r['nm']:,} nm · ${r['cpt']:,.2f}/т")
            fig.add_trace(go.Scattergeo(
                lat=lats, lon=lons, mode="lines+markers",
                line=dict(color="#38648a", width=2.5),
                marker=dict(size=[8 if i in (0,len(pts)-1) else 4
                                  for i in range(len(pts))],
                            color="#15202b", symbol="circle"),
                name=r.get("shipment_id",""),
                hovertemplate=label+"<extra></extra>",
            ))

    # все порты
    port_lats  = [v[0] for v in NODES.values() if v[2]=="p"]
    port_lons  = [v[1] for v in NODES.values() if v[2]=="p"]
    port_names = [k    for k,v in NODES.items() if v[2]=="p"]
    fig.add_trace(go.Scattergeo(
        lat=port_lats, lon=port_lons, mode="markers",
        marker=dict(size=5, color="#38648a",
                    line=dict(color="#15202b", width=0.8)),
        text=port_names,
        hovertemplate="<b>%{text}</b><extra></extra>",
        name="Порты", showlegend=True,
    ))

    # проходные точки
    jp_lats  = [v[0] for v in NODES.values() if v[2]=="j"]
    jp_lons  = [v[1] for v in NODES.values() if v[2]=="j"]
    jp_names = [k    for k,v in NODES.items() if v[2]=="j"]
    fig.add_trace(go.Scattergeo(
        lat=jp_lats, lon=jp_lons, mode="markers",
        marker=dict(size=4, color="#c2d8e6",
                    symbol="diamond",
                    line=dict(color="#15202b", width=0.8)),
        text=jp_names,
        hovertemplate="<b>%{text}</b><extra></extra>",
        name="Waypoints", showlegend=True,
    ))

    fig.update_layout(
        height=620, margin=dict(l=0,r=0,t=0,b=0),
        paper_bgcolor="#e9edf1",
        geo=dict(
            showland=True,  landcolor="#dfe7ee",
            showocean=True, oceancolor="#c2d8e6",
            showcoastlines=True, coastlinecolor="#b9c7d3", coastlinewidth=0.8,
            showlakes=True, lakecolor="#c2d8e6",
            showcountries=True, countrycolor="#b9c7d3", countrywidth=0.5,
            showframe=True, framecolor="#15202b", framewidth=1,
            projection_type="natural earth",
            center=dict(lat=rb["lat"][0], lon=rb["lon"][0]),
            projection_scale=rb["zoom"],
        ),
        legend=dict(orientation="h", y=-0.04, x=0,
                    font=dict(size=10, family=F, color="#15202b"),
                    bgcolor="rgba(255,255,255,0.85)",
                    bordercolor="#15202b", borderwidth=1),
        font=dict(family=F),
        uirevision="map",
    )
    return fig

# ==========================================================================
#  DASH APP
# ==========================================================================
app = Dash(__name__, suppress_callback_exceptions=True)
app.title = "Edible Oils Charterer Planner"

port_list   = sorted(k for k,v in NODES.items() if v[2]=="p")
port_opts   = [{"label":k,"value":k} for k in port_list]
vessel_name_opts = lambda vl: [{"label":v["name"],"value":v["name"]} for v in vl]

def tab_s(sel=False):
    b = {"fontWeight":"600","fontSize":"12px","padding":"6px 12px"}
    return {**b,"fontWeight":"800","color":C["blue"],
            "borderTop":f"2px solid {C['blue']}"} if sel else b

app.layout = html.Div([
    html.Div([
        html.Span("EDIBLE OILS PLANNER",
            style={"fontWeight":"800","fontSize":"14px",
                   "letterSpacing":"0.07em","color":"#fff"}),
        html.Span("  ·  charterer  ·  cost & $/t  ·  fleet  ·  routes  ·  tender",
            style={"fontSize":"11px","color":C["blue_l"],"marginLeft":"8px"}),
    ], style={"padding":"9px 16px","background":C["ink"],
              "borderBottom":f"2px solid {C['blue']}","display":"flex",
              "alignItems":"center"}),

    dcc.Tabs(id="main-tabs", value="tab-refs",
        colors={"border":C["ink"],"primary":C["blue"],"background":C["bg"]},
        style={"fontFamily":F},
        children=[
            dcc.Tab(label="Справочники",   value="tab-refs",
                style=tab_s(), selected_style=tab_s(True)),
            dcc.Tab(label="Отправки",      value="tab-shipments",
                style=tab_s(), selected_style=tab_s(True)),
            dcc.Tab(label="Маршруты / Линии", value="tab-lines",
                style=tab_s(), selected_style=tab_s(True)),
            dcc.Tab(label="Тендер",        value="tab-tender",
                style=tab_s(), selected_style=tab_s(True)),
            dcc.Tab(label="Расписание",    value="tab-schedule",
                style=tab_s(), selected_style=tab_s(True)),
            dcc.Tab(label="Годовой план",  value="tab-plan",
                style=tab_s(), selected_style=tab_s(True)),
            dcc.Tab(label="Спот оценка",   value="tab-spot",
                style=tab_s(), selected_style=tab_s(True)),
            dcc.Tab(label="Карта",         value="tab-map",
                style=tab_s(), selected_style=tab_s(True)),
            dcc.Tab(label="Сценарии",      value="tab-scenarios",
                style=tab_s(), selected_style=tab_s(True)),
            dcc.Tab(label="Экспорт",       value="tab-export",
                style=tab_s(), selected_style=tab_s(True)),
        ]),

    html.Div(id="tab-content", style={"padding":"10px 13px","overflowY":"auto",
        "flex":"1","background":C["bg"]}),

    dcc.Store(id="s-cargoes",   data=DEFAULT_CARGOES),
    dcc.Store(id="s-bunker",    data=DEFAULT_BUNKER),
    dcc.Store(id="s-vessels",   data=DEFAULT_VESSELS),
    dcc.Store(id="s-routes",    data=DEFAULT_ROUTES),
    dcc.Store(id="s-shipments", data=[]),
    dcc.Store(id="s-calc",      data={}),
    dcc.Store(id="s-scenarios", data=[]),
    dcc.Store(id="s-snap",      data=None),

    dcc.Download(id="dl-xlsx"),
    dcc.Download(id="dl-json"),
    dcc.Download(id="dl-scenario"),
    dcc.Download(id="dl-scenarios-all"),
    dcc.Download(id="dl-template"),
], style={"display":"flex","flexDirection":"column","height":"100vh",
          "fontFamily":F,"color":C["ink"],"background":C["bg"]})

# ==========================================================================
#  TAB RENDERER
# ==========================================================================
@app.callback(
    Output("tab-content","children"),
    Input("main-tabs","value"),
    State("s-cargoes","data"), State("s-bunker","data"),
    State("s-vessels","data"), State("s-routes","data"),
    State("s-shipments","data"), State("s-calc","data"),
    State("s-scenarios","data"),
)
def render_tab(tab, cargoes, bunker, vessels, routes, shipments, calc, scenarios):
    cargoes = cargoes or DEFAULT_CARGOES
    bunker  = bunker  or DEFAULT_BUNKER
    vessels = vessels or DEFAULT_VESSELS
    routes  = routes  or DEFAULT_ROUTES

    # ── СПРАВОЧНИКИ ───────────────────────────────────────────────────────
    if tab == "tab-refs":
        def compact_tbl(data, columns, tbl_id):
            return dash_table.DataTable(
                id=tbl_id, data=data, columns=columns,
                editable=True, row_deletable=True,
                row_addable_text="+ строка",
                **TBL,
                style_cell={**TBL["style_cell"],"height":"22px","fontSize":"10.5px"},
            )

        cargo_cols = [
            {"name":"Код","id":"code","editable":True},
            {"name":"Название","id":"name","editable":True},
            {"name":"Группа","id":"group","editable":True},
            {"name":"Плотн.","id":"density","editable":True,"type":"numeric"},
            {"name":"Тепл.°C","id":"heat_deg","editable":True,"type":"numeric"},
            {"name":"mt/d","id":"heat_mt","editable":True,"type":"numeric"},
            {"name":"Мин.партия","id":"parcel_min","editable":True,"type":"numeric"},
            {"name":"Макс.партия","id":"parcel_max","editable":True,"type":"numeric"},
            {"name":"Танки","id":"tanks","editable":True},
            {"name":"Заметки","id":"notes","editable":True},
        ]
        bunker_cols = [
            {"name":"Порт","id":"port","editable":True},
            {"name":"VLSFO","id":"vlsfo","editable":True,"type":"numeric"},
            {"name":"MGO","id":"mgo","editable":True,"type":"numeric"},
            {"name":"Port dues","id":"port_dues","editable":True,"type":"numeric"},
            {"name":"Canal dues","id":"canal_dues","editable":True,"type":"numeric"},
            {"name":"Agents","id":"agents","editable":True,"type":"numeric"},
        ]
        vessel_cols = [
            {"name":"Судно","id":"name","editable":True},
            {"name":"Класс","id":"cls","editable":True},
            {"name":"Intake mt","id":"intake","editable":True,"type":"numeric"},
            {"name":"Уз","id":"speed","editable":True,"type":"numeric"},
            {"name":"VLSFO/d","id":"cons_sea","editable":True,"type":"numeric"},
            {"name":"MGO/d","id":"cons_mgo","editable":True,"type":"numeric"},
            {"name":"Порт/d","id":"cons_port","editable":True,"type":"numeric"},
            {"name":"Bunk.tank","id":"tank","editable":True,"type":"numeric"},
            {"name":"Load.h","id":"load_hours","editable":True,"type":"numeric"},
            {"name":"Disch.h","id":"disch_hours","editable":True,"type":"numeric"},
            {"name":"TC $/d","id":"tc_rate","editable":True,"type":"numeric"},
            {"name":"Opex $/d","id":"opex_day","editable":True,"type":"numeric"},
            {"name":"Стоим.$","id":"value","editable":True,"type":"numeric"},
            {"name":"Оффхайр","id":"offhire_days","editable":True,"type":"numeric"},
            {"name":"Heat mt/d","id":"heat_mt_day","editable":True,"type":"numeric"},
            {"name":"Статус","id":"status","editable":True},
            {"name":"Заметки","id":"notes","editable":True},
        ]
        route_cols = [
            {"name":"ID","id":"id","editable":True},
            {"name":"Название","id":"name","editable":True},
            {"name":"Порт погрузки","id":"load","editable":True,
             "presentation":"dropdown"},
            {"name":"Порт выгрузки","id":"disch","editable":True,
             "presentation":"dropdown"},
            {"name":"Груз по умолч.","id":"cargo","editable":True},
            {"name":"Заметки","id":"notes","editable":True},
        ]
        up_style = {"padding":"10px","border":f"1.5px dashed {C['ink_l']}",
                    "borderRadius":"3px","textAlign":"center","fontSize":"10.5px",
                    "color":C["muted"],"cursor":"pointer","marginBottom":"6px"}
        return html.Div([
            html.Div([
                html.Div([
                    card([sec("Грузы"),
                          compact_tbl(cargoes, cargo_cols, "tbl-cargoes")]),
                    card([sec("Суда"),
                          html.Button("+ Добавить судно", id="btn-add-vessel",
                              style={**btng,"width":"auto","marginBottom":"8px",
                                     "marginTop":"0","fontSize":"11px"}),
                          compact_tbl(vessels, vessel_cols, "tbl-vessels")]),
                ], style={"flex":"1","minWidth":"0","marginRight":"8px"}),
                html.Div([
                    card([sec("Бункер и портовые расходы"),
                          compact_tbl(bunker, bunker_cols, "tbl-bunker")]),
                    card([sec("Стандартные маршруты"),
                          compact_tbl(routes, route_cols, "tbl-routes",),]),
                ], style={"flex":"1","minWidth":"0"}),
            ], style={"display":"flex","gap":"0","alignItems":"flex-start"}),
            card([
                sec("Загрузка из Excel"),
                html.Div([
                    html.Div([
                        html.Div("Грузы (лист Грузы)",
                            style={"fontSize":"10.5px","color":C["muted"],"marginBottom":"3px"}),
                        dcc.Upload(id="up-cargoes", children=html.Div("XLSX", style=up_style)),
                    ], style={"flex":"1"}),
                    html.Div([
                        html.Div("Бункер (лист Бункер)",
                            style={"fontSize":"10.5px","color":C["muted"],"marginBottom":"3px"}),
                        dcc.Upload(id="up-bunker", children=html.Div("XLSX", style=up_style)),
                    ], style={"flex":"1"}),
                    html.Div([
                        html.Div("Суда (лист Суда)",
                            style={"fontSize":"10.5px","color":C["muted"],"marginBottom":"3px"}),
                        dcc.Upload(id="up-vessels", children=html.Div("XLSX", style=up_style)),
                    ], style={"flex":"1"}),
                ], style={"display":"flex","gap":"10px"}),
                html.Button("Скачать шаблон Excel", id="btn-template",
                    style={**btng,"width":"auto","marginTop":"4px","fontSize":"11px"}),
            ]),
        ])

    # ── ОТПРАВКИ ──────────────────────────────────────────────────────────
    elif tab == "tab-shipments":
        c_opts = [{"label":f"{c['code']} {c['name']}","value":c["code"]} for c in cargoes]
        v_opts = vessel_name_opts(vessels)
        r_opts = [{"label":f"{r['id']} {r['name']}","value":r["id"]} for r in routes]
        shp_cols = [
            {"name":"ID","id":"id","editable":True},
            {"name":"Линия/Маршрут","id":"route_id","editable":True,"presentation":"dropdown"},
            {"name":"Груз","id":"cargo_code","editable":True,"presentation":"dropdown"},
            {"name":"Порт погрузки","id":"load_port","editable":True,"presentation":"dropdown"},
            {"name":"Порт выгрузки","id":"disch_port","editable":True,"presentation":"dropdown"},
            {"name":"Кол-во mt","id":"quantity","editable":True,"type":"numeric"},
            {"name":"Судно","id":"vessel_id","editable":True,"presentation":"dropdown"},
            {"name":"Контракт","id":"contract_type","editable":True,"presentation":"dropdown"},
            {"name":"Спот $/d","id":"spot_rate","editable":True,"type":"numeric"},
            {"name":"Месяц","id":"month","editable":True,"type":"numeric"},
            {"name":"Дата нач.","id":"date_start","editable":True},
            {"name":"Heat mt/d","id":"heat_mt_day","editable":True,"type":"numeric"},
            {"name":"Статус","id":"status","editable":True},
        ]
        dropdowns = {
            "cargo_code":    {"options":c_opts},
            "load_port":     {"options":port_opts},
            "disch_port":    {"options":port_opts},
            "vessel_id":     {"options":v_opts},
            "contract_type": {"options":[{"label":t,"value":t} for t in CONTRACT_TYPES]},
            "route_id":      {"options":r_opts},
        }
        seed = shipments if shipments else [
            {"id":f"S{i+1:03d}","route_id":"R01","cargo_code":"CPO",
             "load_port":"Pasir Gudang","disch_port":"Rotterdam",
             "quantity":20000,"vessel_id":vessels[0]["name"],
             "contract_type":"TC In","spot_rate":14000,"month":3,
             "date_start":"2026-03-01","heat_mt_day":1.8,"status":"planned"}
            for i in range(4)
        ]
        return html.Div([
            card([
                sec("Грузовые отправки"),
                html.P("Базовая единица планирования. Привяжите к стандартному маршруту, "
                       "выберите судно и контракт, задайте дату начала.",
                    style={"fontSize":"10.5px","color":C["muted"],"marginBottom":"8px"}),
                html.Div([
                    html.Button("Рассчитать затраты", id="btn-calc-all",
                        style={**btn,"width":"auto","marginRight":"8px","marginTop":"0"}),
                    html.Button("+ Добавить отправку", id="btn-add-shp",
                        style={**btng,"width":"auto","marginTop":"0"}),
                ], style={"display":"flex","marginBottom":"8px","gap":"6px"}),
            ]),
            card([
                dash_table.DataTable(id="tbl-shipments", data=seed,
                    columns=shp_cols, editable=True, row_deletable=True,
                    dropdown=dropdowns, **TBL),
            ]),
            html.Div(id="shp-results"),
        ])

    # ── МАРШРУТЫ / ЛИНИИ ──────────────────────────────────────────────────
    elif tab == "tab-lines":
        shp_res = (calc or {}).get("shipments",[])
        if not shp_res:
            return card([sec("Маршруты / Линии"),
                html.P("Рассчитайте отправки на вкладке «Отправки».",
                    style={"color":C["muted"]})])

        by_route = {}
        for r in shp_res:
            by_route.setdefault(r.get("route_id","—"),[]).append(r)

        cards = []
        for rid, items in sorted(by_route.items()):
            rdef = next((x for x in routes if x["id"]==rid), {})
            tot_cost = sum(x["total"] for x in items)
            tot_qty  = sum(x.get("qty",0) for x in items)
            avg_cpt  = tot_cost/tot_qty if tot_qty else 0
            avg_days = sum(x["total_days"] for x in items)/len(items)

            fig = go.Figure(go.Bar(
                x=MONTH_NAMES,
                y=[sum(x["total"] for x in items
                       if int(x.get("month",0))==i+1) for i in range(12)],
                marker_color=C["blue_l"],
                marker_line_color=C["ink"], marker_line_width=1,
            ))
            fig.update_layout(height=120,
                margin=dict(l=30,r=8,t=6,b=24),
                paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
                font=dict(family=F,size=9,color=C["ink"]),
                showlegend=False,
                xaxis=dict(showgrid=False),
                yaxis=dict(gridcolor=C["ink_l"],tickprefix="$"))

            tbl = dash_table.DataTable(
                data=[{
                    "ID":     x.get("shipment_id",""),
                    "Судно":  x.get("vessel",""),
                    "Контракт": x.get("contract",""),
                    "mt":     f"{x.get('qty',0):,.0f}",
                    "Дни":    x["total_days"],
                    "Бункер $": money(x["bunker"]),
                    "Порты $":  money(x["ports"]),
                    "Найм $":   money(x["hire"]),
                    "Риск $":   money(x["risk"]),
                    "Итого $":  money(x["total"]),
                    "$/т":      f"${x['cpt']:,.2f}",
                    "SECA":   ", ".join(x["seca"]) or "—",
                } for x in items],
                columns=[{"name":c,"id":c} for c in
                    ["ID","Судно","Контракт","mt","Дни",
                     "Бункер $","Порты $","Найм $","Риск $","Итого $","$/т","SECA"]],
                **TBL)

            cards.append(card([
                html.Div([
                    html.Div(f"{rdef.get('name','Линия ')+' · '+rid}",
                        style={"fontSize":"13px","fontWeight":"800","color":C["ink"]}),
                    html.Div(
                        f"{rdef.get('load','')} → {rdef.get('disch','')}  ·  "
                        f"{len(items)} отпр.  ·  {tot_qty:,.0f} mt",
                        style={"fontSize":"10.5px","color":C["muted"]}),
                ], style={"marginBottom":"8px"}),
                html.Div([
                    kpi("Всего затрат",   money(tot_cost), f"{tot_qty:,.0f} mt"),
                    kpi("Ср. $/т",        f"${avg_cpt:,.2f}", "cost per tonne"),
                    kpi("Ср. рейс, дн.",  f"{avg_days:.1f}", "total voyage"),
                ], style={"display":"flex","gap":"8px","flexWrap":"wrap",
                          "marginBottom":"10px"}),
                dcc.Graph(figure=fig, config={"displayModeBar":False}),
                html.Div(style={"height":"6px"}),
                tbl,
            ]))
        return html.Div(cards)

    # ── ТЕНДЕР — матрица судно × маршрут ──────────────────────────────────
    elif tab == "tab-tender":
        return html.Div([
            card([
                sec("Тендер: сравнение судов на маршруте"),
                html.P("Выберите маршрут (порт погрузки → выгрузки), груз и объём. "
                       "Система рассчитает затраты $/т для каждого судна.",
                    style={"fontSize":"10.5px","color":C["muted"],"marginBottom":"8px"}),
                html.Div([
                    html.Div([
                        html.Div("Порт погрузки", style={"fontSize":"10.5px","color":C["muted"],
                            "marginBottom":"3px","fontWeight":"700"}),
                        dcc.Dropdown(id="td-load", options=port_opts,
                            value="Pasir Gudang", clearable=False, style={"fontSize":"12px"}),
                    ], style={"flex":"1"}),
                    html.Div([
                        html.Div("Порт выгрузки", style={"fontSize":"10.5px","color":C["muted"],
                            "marginBottom":"3px","fontWeight":"700"}),
                        dcc.Dropdown(id="td-disch", options=port_opts,
                            value="Rotterdam", clearable=False, style={"fontSize":"12px"}),
                    ], style={"flex":"1"}),
                    html.Div([
                        html.Div("Кол-во mt", style={"fontSize":"10.5px","color":C["muted"],
                            "marginBottom":"3px","fontWeight":"700"}),
                        dcc.Input(id="td-qty", type="number", value=20000, style=inp),
                    ], style={"flex":"1"}),
                    html.Div([
                        html.Div("Месяц", style={"fontSize":"10.5px","color":C["muted"],
                            "marginBottom":"3px","fontWeight":"700"}),
                        dcc.Dropdown(id="td-month",
                            options=[{"label":"Any","value":0}]+
                                    [{"label":m,"value":i+1} for i,m in enumerate(MONTH_NAMES)],
                            value=0, clearable=False, style={"fontSize":"12px"}),
                    ], style={"flex":"1"}),
                    html.Div([
                        html.Div("Контракт", style={"fontSize":"10.5px","color":C["muted"],
                            "marginBottom":"3px","fontWeight":"700"}),
                        dcc.Dropdown(id="td-contract",
                            options=[{"label":t,"value":t} for t in CONTRACT_TYPES],
                            value="TC In", clearable=False, style={"fontSize":"12px"}),
                    ], style={"flex":"1"}),
                    html.Div([
                        html.Div("Heat mt/d", style={"fontSize":"10.5px","color":C["muted"],
                            "marginBottom":"3px","fontWeight":"700"}),
                        dcc.Input(id="td-heat", type="number", value=1.8, step=0.1, style=inp),
                    ], style={"flex":"0.5"}),
                ], style={"display":"flex","gap":"10px","flexWrap":"wrap","marginBottom":"10px"}),
                html.Button("Сравнить суда", id="btn-tender",
                    style={**btn,"width":"auto","marginTop":"0"}),
            ]),
            html.Div(id="tender-results"),
        ])

    # ── РАСПИСАНИЕ ────────────────────────────────────────────────────────
    elif tab == "tab-schedule":
        shp_res = (calc or {}).get("shipments",[])
        v_opts = vessel_name_opts(vessels)
        return html.Div([
            card([
                sec("Расписание по судну"),
                html.Div([
                    html.Div([
                        html.Div("Судно", style={"fontSize":"10.5px","color":C["muted"],
                            "marginBottom":"3px","fontWeight":"700"}),
                        dcc.Dropdown(id="sch-vessel",
                            options=v_opts,
                            value=vessels[0]["name"] if vessels else None,
                            clearable=False, style={"fontSize":"12px"}),
                    ], style={"flex":"1"}),
                    html.Div([
                        html.Div("Дата начала", style={"fontSize":"10.5px","color":C["muted"],
                            "marginBottom":"3px","fontWeight":"700"}),
                        dcc.Input(id="sch-start", type="text",
                            value="2026-01-01", placeholder="YYYY-MM-DD", style=inp),
                    ], style={"flex":"1"}),
                ], style={"display":"flex","gap":"10px","marginBottom":"10px"}),
                html.Button("Построить расписание", id="btn-schedule",
                    style={**btn,"width":"auto","marginTop":"0"}),
            ]),
            html.Div(id="schedule-results"),
        ])

    # ── ГОДОВОЙ ПЛАН ──────────────────────────────────────────────────────
    elif tab == "tab-plan":
        shp_res = (calc or {}).get("shipments",[])
        if not shp_res:
            return card([sec("Годовой план"),
                html.P("Рассчитайте отправки.", style={"color":C["muted"]})])

        cost_m  = [0.]*12; bunk_m  = [0.]*12
        hire_m  = [0.]*12; risk_m  = [0.]*12
        qty_m   = [0.]*12; days_m  = [0.]*12; cpt_sum = [0.]*12; cpt_n = [0]*12

        for r in shp_res:
            m = int(r.get("month") or 0)
            idxs = [m-1] if 1<=m<=12 else list(range(12))
            w = 1. if len(idxs)==1 else 1./12
            for i in idxs:
                cost_m[i]  += r["total"]      * w
                bunk_m[i]  += r["bunker"]     * w
                hire_m[i]  += r["hire"]       * w
                risk_m[i]  += r["risk"]       * w
                qty_m[i]   += r.get("qty",0)  * w
                days_m[i]  += r["total_days"] * w
                cpt_sum[i] += r["cpt"]        * w
                cpt_n[i]   += w

        total_cost = sum(cost_m)
        total_qty  = sum(qty_m)
        avg_cpt    = total_cost/total_qty if total_qty else 0

        # waterfall: бункер / найм / порты / риск → итого
        tot_bunk = sum(r["bunker"] for r in shp_res)
        tot_hire = sum(r["hire"]   for r in shp_res)
        tot_port = sum(r["ports"]  for r in shp_res)
        tot_risk = sum(r["risk"]   for r in shp_res)
        wf = go.Figure(go.Waterfall(
            orientation="v",
            measure=["absolute","relative","relative","relative","total"],
            x=["Бункер","Найм/Opex","Порты/Каналы","Риск AWRP","Итого"],
            y=[tot_bunk, tot_hire, tot_port, tot_risk, total_cost],
            connector={"line":{"color":C["ink"],"width":1}},
            increasing={"marker":{"color":C["blue_l"],
                "line":{"color":C["ink"],"width":1}}},
            decreasing={"marker":{"color":C["red_l"],
                "line":{"color":C["ink"],"width":1}}},
            totals={"marker":{"color":C["surf2"],
                "line":{"color":C["ink"],"width":1.5}}},
            text=[f"${v/1e6:.1f}M" for v in
                  [tot_bunk,tot_hire,tot_port,tot_risk,total_cost]],
            textposition="outside",
        ))
        wf.update_layout(height=260, margin=dict(l=40,r=10,t=20,b=30),
            paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
            font=dict(family=F,size=10,color=C["ink"]),
            yaxis=dict(gridcolor=C["ink_l"],tickprefix="$"),
            xaxis=dict(showgrid=False))

        # помесячные затраты
        mb = go.Figure()
        mb.add_trace(go.Bar(name="Бункер",   x=MONTH_NAMES, y=bunk_m,
            marker_color=C["blue_l"], marker_line_color=C["ink"],marker_line_width=1))
        mb.add_trace(go.Bar(name="Найм",     x=MONTH_NAMES, y=hire_m,
            marker_color=C["blue"],   marker_line_color=C["ink"],marker_line_width=1))
        mb.add_trace(go.Bar(name="Порты/Риск",x=MONTH_NAMES,
            y=[risk_m[i]+cost_m[i]-bunk_m[i]-hire_m[i] for i in range(12)],
            marker_color=C["ink_l"],  marker_line_color=C["ink"],marker_line_width=1))
        mb.add_trace(go.Scatter(name="$/т ср.", x=MONTH_NAMES,
            y=[cpt_sum[i]/cpt_n[i] if cpt_n[i] else 0 for i in range(12)],
            mode="lines+markers", yaxis="y2",
            line=dict(color=C["red"],width=2),
            marker=dict(size=5,color=C["red"],line=dict(color=C["ink"],width=1))))
        mb.update_layout(barmode="stack", height=260,
            margin=dict(l=40,r=60,t=10,b=30),
            paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
            font=dict(family=F,size=10,color=C["ink"]),
            legend=dict(orientation="h",y=-0.3),
            xaxis=dict(showgrid=False),
            yaxis=dict(gridcolor=C["ink_l"],tickprefix="$"),
            yaxis2=dict(overlaying="y",side="right",
                tickprefix="$",title="$/т",showgrid=False))

        bud = [{"Месяц":MONTH_NAMES[i],
                "Затраты $":    f"{cost_m[i]:,.0f}",
                "Бункер $":     f"{bunk_m[i]:,.0f}",
                "Найм $":       f"{hire_m[i]:,.0f}",
                "Риск $":       f"{risk_m[i]:,.0f}",
                "$/т":          f"{cpt_sum[i]/cpt_n[i]:.2f}" if cpt_n[i] else "—",
                "Объём mt":     f"{qty_m[i]:,.0f}",
                "Судо-дни":     f"{days_m[i]:.1f}"}
               for i in range(12)]
        bud.append({"Месяц":"ИТОГО",
            "Затраты $":f"{total_cost:,.0f}", "Бункер $":f"{sum(bunk_m):,.0f}",
            "Найм $":f"{sum(hire_m):,.0f}", "Риск $":f"{sum(risk_m):,.0f}",
            "$/т":f"{avg_cpt:.2f}", "Объём mt":f"{total_qty:,.0f}",
            "Судо-дни":f"{sum(days_m):.1f}"})

        return html.Div([
            html.Div([
                kpi("Годовые затраты",  money(total_cost)),
                kpi("Объём",           f"{total_qty/1000:,.0f} Kmt",
                    f"{len(shp_res)} отправок"),
                kpi("Ср. $/т",         f"${avg_cpt:,.2f}", "cost per tonne"),
                kpi("Бункер",          f"{sum(bunk_m)/total_cost*100:.0f}% затрат"
                    if total_cost else "—", money(sum(bunk_m))),
            ], style={"display":"flex","gap":"8px","flexWrap":"wrap","marginBottom":"10px"}),
            card([sec("Структура затрат"),
                  dcc.Graph(figure=wf, config={"displayModeBar":False})]),
            card([sec("Помесячно: затраты + $/т"),
                  dcc.Graph(figure=mb, config={"displayModeBar":False})]),
            card([sec("Бюджет по месяцам"),
                  dash_table.DataTable(data=bud,
                      columns=[{"name":c,"id":c} for c in bud[0].keys()],
                      **TBL)]),
        ])

    # ── СПОТ ОЦЕНКА ───────────────────────────────────────────────────────
    elif tab == "tab-spot":
        shp_res = (calc or {}).get("shipments",[])
        if not shp_res:
            return card([sec("Спот оценка"),
                html.P("Рассчитайте отправки.", style={"color":C["muted"]})])

        rows, own_cpt, spot_cpt, labels_s, savings = [], [], [], [], []
        for r in shp_res:
            sr = float(next((s.get("spot_rate") or 0
                for s in (shipments or []) if s.get("id")==r.get("shipment_id")),0))
            if sr <= 0: continue
            qty = r.get("qty",0)
            if qty <= 0: continue
            spot_total = sr*r["total_days"] + r["bunker"] + r["ports"] + r["risk"]
            spot_cpt_v = spot_total/qty
            saving_pmt = r["cpt"] - spot_cpt_v   # + = своё дешевле
            own_cpt.append(r["cpt"])
            spot_cpt.append(spot_cpt_v)
            labels_s.append(f"{r['load'][:10]}→{r['disch'][:10]}")
            savings.append(saving_pmt)
            rows.append({
                "ID":             r.get("shipment_id",""),
                "Маршрут":        f"{r['load']} → {r['disch']}",
                "mt":             f"{qty:,.0f}",
                "TC/Контракт":    r.get("contract",""),
                "$/т (своё)":     f"${r['cpt']:,.2f}",
                "$/т (спот)":     f"${spot_cpt_v:,.2f}",
                "Разница $/т":    f"${saving_pmt:,.2f}",
                "Итого своё $":   money(r["total"]),
                "Итого спот $":   money(spot_total),
                "Решение":        "Своё дешевле" if saving_pmt >= 0 else "Спот дешевле",
            })

        if not rows:
            return card([sec("Спот оценка"),
                html.P("Укажите «Спот $/d» в отправках и пересчитайте.",
                    style={"color":C["muted"]})])

        # scatter: $/т своё vs $/т спот
        mn_ = min(own_cpt+spot_cpt)*0.95; mx_ = max(own_cpt+spot_cpt)*1.05
        sc = go.Figure()
        sc.add_shape(type="line",x0=mn_,y0=mn_,x1=mx_,y1=mx_,
            line=dict(color=C["ink_l"],width=1,dash="dash"))
        sc.add_trace(go.Scatter(x=spot_cpt, y=own_cpt,
            mode="markers+text", text=labels_s, textposition="top center",
            textfont=dict(size=9,color=C["ink"]),
            marker=dict(size=11,color=C["blue_l"],
                line=dict(color=C["ink"],width=1.5)),
            hovertemplate="Спот: $%{x:.2f}/т<br>Своё: $%{y:.2f}/т<extra></extra>"))
        sc.update_layout(height=300,
            margin=dict(l=50,r=10,t=10,b=40),
            paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
            font=dict(family=F,size=10,color=C["ink"]),
            xaxis=dict(title="$/т спот",gridcolor=C["ink_l"]),
            yaxis=dict(title="$/т своё",gridcolor=C["ink_l"]))

        colors_s = [C["green_l"] if s >= 0 else C["red_l"] for s in savings]
        sv = go.Figure(go.Bar(
            x=[r["ID"] for r in rows], y=savings,
            marker_color=colors_s,
            marker_line_color=C["ink"], marker_line_width=1,
            text=[f"${s:,.2f}/т" for s in savings],
            textposition="outside",
            hovertemplate="Разница: $%{y:.2f}/т<extra></extra>"))
        sv.update_layout(height=220,
            margin=dict(l=40,r=10,t=10,b=40),
            paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
            font=dict(family=F,size=10,color=C["ink"]),
            xaxis=dict(title="Отправка",showgrid=False),
            yaxis=dict(title="$/т разница",gridcolor=C["ink_l"],
                zeroline=True,zerolinecolor=C["ink"],zerolinewidth=1.5))

        return html.Div([
            card([sec("$/т своё vs $/т спот — оценка стоимости тонны"),
                  html.P("Точки ниже диагонали: спот дешевле. "
                         "Точки выше: своё судно / TC выгоднее.",
                      style={"fontSize":"10.5px","color":C["muted"],"marginBottom":"6px"}),
                  dcc.Graph(figure=sc, config={"displayModeBar":False})]),
            card([sec("Разница $/т (+ своё дешевле, − спот дешевле)"),
                  dcc.Graph(figure=sv, config={"displayModeBar":False})]),
            card([dash_table.DataTable(data=rows,
                columns=[{"name":c,"id":c} for c in rows[0].keys()],
                **TBL)]),
        ])

    # ── КАРТА ─────────────────────────────────────────────────────────────
    elif tab == "tab-map":
        shp_res = (calc or {}).get("shipments",[])
        region_opts = [{"label":k,"value":k} for k in REGION_BOUNDS]
        return html.Div([
            card([
                sec("Карта маршрутов"),
                html.Div([
                    html.Div("Регион / Зум:", style={"fontSize":"10.5px","color":C["muted"],
                        "fontWeight":"700","marginRight":"8px","lineHeight":"32px"}),
                    dcc.Dropdown(id="map-region", options=region_opts,
                        value="Мир целиком", clearable=False,
                        style={"fontSize":"12px","width":"220px"}),
                    html.Button("Обновить карту", id="btn-map-refresh",
                        style={**btng,"width":"auto","marginTop":"0",
                               "marginLeft":"10px","fontSize":"11px"}),
                ], style={"display":"flex","alignItems":"center","marginBottom":"8px",
                          "gap":"0","flexWrap":"wrap"}),
                html.P("Зелёные рамки — SECA ECA. Красные — зоны военного риска AWRP/LOH. "
                       "Синие линии — рассчитанные маршруты отправок.",
                    style={"fontSize":"10.5px","color":C["muted"],"marginBottom":"6px"}),
                dcc.Graph(id="main-map",
                    figure=build_map(shp_res, "Мир целиком"),
                    config={"displayModeBar":True,
                            "modeBarButtonsToAdd":["select2d","lasso2d"],
                            "scrollZoom":True}),
            ]),
        ])

    # ── СЦЕНАРИИ ──────────────────────────────────────────────────────────
    elif tab == "tab-scenarios":
        scen = scenarios or []
        rows = [html.Div([
            html.Div([
                html.Div(s.get("name","—"), style={"fontSize":"12px",
                    "fontWeight":"700","color":C["ink"]}),
                html.Div(f"{s.get('saved','')} · "
                         f"{len(s.get('shipments',[]))} отпр. · "
                         f"{len(s.get('vessels',[]))} суд.",
                    style={"fontSize":"10.5px","color":C["muted"]}),
            ], style={"flex":"1"}),
            html.Button("Загрузить", id={"type":"sc-load","index":i}, n_clicks=0,
                style={**btng,"width":"auto","marginTop":"0","padding":"4px 9px",
                       "fontSize":"11px"}),
            html.Button("JSON", id={"type":"sc-exp","index":i}, n_clicks=0,
                style={**btng,"width":"auto","marginTop":"0","padding":"4px 9px",
                       "fontSize":"11px"}),
            html.Button("✕", id={"type":"sc-del","index":i}, n_clicks=0,
                style={**btn,"width":"auto","marginTop":"0","padding":"4px 9px",
                       "fontSize":"11px","background":C["red"]}),
        ], style={"display":"flex","gap":"6px","alignItems":"center",
                  "padding":"6px 0","borderBottom":f"1px solid {C['ink_l']}"})]
            for i, s in enumerate(scen)]

        return html.Div([
            card([
                sec("Сохранить снимок"),
                dcc.Input(id="sc-name", type="text",
                    placeholder="Имя сценария, например Base Q1 2026", style=inp),
                html.Button("Сохранить", id="btn-save-sc",
                    style={**btn,"width":"auto","marginTop":"6px"}),
            ]),
            card([
                sec("Импорт из JSON"),
                dcc.Upload(id="up-scenario", multiple=False,
                    children=html.Div("Перетащите JSON сценария или нажмите",
                        style={"padding":"10px","border":f"1.5px dashed {C['ink_l']}",
                               "borderRadius":"3px","textAlign":"center","fontSize":"10.5px",
                               "color":C["muted"],"cursor":"pointer"})),
                html.Button("Экспорт всех сценариев", id="btn-exp-all",
                    style={**btng,"width":"auto","marginTop":"6px","fontSize":"11px"}),
            ]),
            card([sec("Сохранённые сценарии")] +
                 (rows if rows else [html.P("Нет сохранённых сценариев.",
                     style={"fontSize":"11px","color":C["muted"]})])),
        ])

    # ── ЭКСПОРТ ───────────────────────────────────────────────────────────
    elif tab == "tab-export":
        shp_res = (calc or {}).get("shipments",[])
        total_cost = sum(r["total"] for r in shp_res)
        total_qty  = sum(r.get("qty",0) for r in shp_res)
        avg_cpt    = total_cost/total_qty if total_qty else 0
        return html.Div([
            card([
                sec("Статус расчёта"),
                html.Div([
                    kpi("Отправок",   str(len(shp_res))),
                    kpi("Затраты",    money(total_cost)),
                    kpi("Ср. $/т",    f"${avg_cpt:,.2f}"),
                ], style={"display":"flex","gap":"8px","flexWrap":"wrap",
                          "marginBottom":"10px"}),
            ]),
            card([
                sec("Скачать"),
                html.P("Нажмите нужную кнопку — файл скачается немедленно.",
                    style={"fontSize":"10.5px","color":C["muted"],"marginBottom":"8px"}),
                html.Button("XLSX — отправки и бюджет", id="btn-dl-xlsx",
                    style={**btn,"marginBottom":"6px"}),
                html.Button("JSON — полный датасет",    id="btn-dl-json",
                    style={**btng,"marginBottom":"6px"}),
            ], style={"maxWidth":"320px"}),
        ])

    return html.Div()

# ==========================================================================
#  STORE SYNC
# ==========================================================================
@app.callback(Output("s-cargoes","data"), Input("tbl-cargoes","data"),
              prevent_initial_call=True)
def sync_cargoes(d): return d or []

@app.callback(Output("s-bunker","data"), Input("tbl-bunker","data"),
              prevent_initial_call=True)
def sync_bunker(d): return d or []

@app.callback(Output("s-vessels","data"), Input("tbl-vessels","data"),
              prevent_initial_call=True)
def sync_vessels(d): return d or []

@app.callback(Output("s-routes","data"), Input("tbl-routes","data"),
              prevent_initial_call=True)
def sync_routes(d): return d or []

@app.callback(Output("s-shipments","data"),
              Input("tbl-shipments","data"),
              Input("btn-add-shp","n_clicks"),
              State("s-shipments","data"),
              State("s-vessels","data"),
              prevent_initial_call=True)
def sync_shipments(tbl, _add, store, vessels):
    if ctx.triggered_id == "btn-add-shp":
        vl = vessels or DEFAULT_VESSELS
        new = {"id":f"S{len(tbl or [])+1:03d}","route_id":"R01",
               "cargo_code":"CPO","load_port":"Pasir Gudang",
               "disch_port":"Rotterdam","quantity":20000,
               "vessel_id":vl[0]["name"],"contract_type":"TC In",
               "spot_rate":14000,"month":0,"date_start":"",
               "heat_mt_day":1.8,"status":"planned"}
        return (tbl or []) + [new]
    return tbl or []

# добавление судна через кнопку
@app.callback(Output("s-vessels","data", allow_duplicate=True),
              Input("btn-add-vessel","n_clicks"),
              State("s-vessels","data"),
              prevent_initial_call=True)
def add_vessel(_n, vessels):
    vl = list(vessels or DEFAULT_VESSELS)
    vl.append({
        "name":f"MT New Vessel {len(vl)+1}","cls":"IMO 2 parcel","intake":20000,
        "speed":13.5,"cons_sea":26,"cons_mgo":19,"cons_port":3.0,"tank":1100,
        "load_hours":30,"disch_hours":36,"status":"TC In","tc_rate":14000,
        "opex_day":0,"value":25000000,"offhire_days":12,"heat_mt_day":1.8,"notes":"",
    })
    return vl

# ==========================================================================
#  ОСНОВНОЙ РАСЧЁТ
# ==========================================================================
def run_calc(shipments, vessels, bunker):
    results, errors = [], []
    for shp in shipments or []:
        vid = shp.get("vessel_id","")
        v   = next((x for x in vessels if x.get("name")==vid),
                   vessels[0] if vessels else None)
        if v is None:
            errors.append(shp.get("id","?")); continue
        # heat_mt_day может быть переопределено в отправке
        vv = dict(v)
        if shp.get("heat_mt_day") is not None:
            vv["heat_mt_day"] = float(shp.get("heat_mt_day") or 0)
        r = voyage_cost(
            load=shp.get("load_port",""),
            disch=shp.get("disch_port",""),
            qty=float(shp.get("quantity") or 0),
            month=int(shp.get("month") or 0),
            vessel=vv,
            bunker_ref=bunker,
            contract=shp.get("contract_type","TC In"),
        )
        if r:
            r["shipment_id"] = shp.get("id","")
            r["route_id"]    = shp.get("route_id","")
            r["vessel"]      = vid
            r["qty"]         = float(shp.get("quantity") or 0)
            r["cargo"]       = shp.get("cargo_code","")
            r["month"]       = int(shp.get("month") or 0)
            r["contract"]    = shp.get("contract_type","")
            r["date_start"]  = shp.get("date_start","")
            results.append(r)
        else:
            errors.append(shp.get("id","?"))
    return results, errors

@app.callback(Output("s-calc","data"),
              Output("shp-results","children"),
              Input("btn-calc-all","n_clicks"),
              State("s-shipments","data"),
              State("s-vessels","data"),
              State("s-bunker","data"),
              prevent_initial_call=True)
def calc_all(n, shipments, vessels, bunker):
    if not shipments: raise PreventUpdate
    results, errors = run_calc(shipments,
        vessels or DEFAULT_VESSELS, bunker or DEFAULT_BUNKER)
    total_cost = sum(r["total"] for r in results)
    total_qty  = sum(r.get("qty",0) for r in results)
    avg_cpt    = total_cost/total_qty if total_qty else 0
    summary = html.Div([
        html.Div([
            kpi("Рассчитано",  str(len(results))),
            kpi("Затраты",     money(total_cost)),
            kpi("Ср. $/т",     f"${avg_cpt:,.2f}"),
        ], style={"display":"flex","gap":"8px","flexWrap":"wrap","marginTop":"8px"}),
        (html.Div(f"Ошибка маршрута: {', '.join(errors)}",
            style={"color":C["red"],"fontSize":"10.5px","marginTop":"5px"})
         if errors else html.Div()),
    ])
    return {"shipments":results}, summary

# ==========================================================================
#  ТЕНДЕР
# ==========================================================================
@app.callback(Output("tender-results","children"),
              Input("btn-tender","n_clicks"),
              State("td-load","value"), State("td-disch","value"),
              State("td-qty","value"), State("td-month","value"),
              State("td-contract","value"), State("td-heat","value"),
              State("s-vessels","data"), State("s-bunker","data"),
              prevent_initial_call=True)
def run_tender(n, load, disch, qty, month, contract, heat, vessels, bunker):
    if not load or not disch: raise PreventUpdate
    vessels = vessels or DEFAULT_VESSELS
    bunker  = bunker  or DEFAULT_BUNKER
    qty_f   = float(qty or 20000)
    heat_f  = float(heat or 0)
    rows = []
    for v in vessels:
        vv = dict(v); vv["heat_mt_day"] = heat_f
        r = voyage_cost(load, disch, qty_f, int(month or 0),
                        vv, bunker, contract)
        if r:
            rows.append({
                "Судно":    v["name"],
                "Класс":    v.get("cls",""),
                "Intake mt":f"{v.get('intake',0):,}",
                "nm":       f"{r['nm']:,}",
                "Дни":      r["total_days"],
                "Бункер $": money(r["bunker"]),
                "Порты $":  money(r["ports"]),
                "Найм $":   money(r["hire"]),
                "Риск $":   money(r["risk"]),
                "Итого $":  money(r["total"]),
                "$/т":      f"${r['cpt']:,.2f}",
                "SECA":     ", ".join(r["seca"]) or "—",
                "Зоны риска": ", ".join(r["risk_zones"]) or "—",
            })

    if not rows:
        return card([html.P("Маршрут не найден.",
            style={"color":C["red"],"fontSize":"11px"})])

    rows.sort(key=lambda x: float(x["$/т"].replace("$","").replace(",","")))
    best = rows[0]["Судно"]

    # бар: $/т по судам
    fig = go.Figure()
    colors_t = [C["green_l"] if r["Судно"]==best else C["blue_l"] for r in rows]
    fig.add_trace(go.Bar(
        x=[r["Судно"] for r in rows],
        y=[float(r["$/т"].replace("$","").replace(",","")) for r in rows],
        marker_color=colors_t,
        marker_line_color=C["ink"], marker_line_width=1,
        text=[r["$/т"] for r in rows], textposition="outside",
        hovertemplate="%{x}<br>$/т: %{y:.2f}<extra></extra>",
    ))
    fig.update_layout(height=220,
        margin=dict(l=40,r=10,t=10,b=60),
        paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
        font=dict(family=F,size=10,color=C["ink"]),
        xaxis=dict(showgrid=False,tickangle=-30),
        yaxis=dict(title="$/т",gridcolor=C["ink_l"],tickprefix="$"))

    # стек затрат
    fig2 = go.Figure()
    for label, key, col in [
        ("Бункер","bunker",C["blue_l"]),("Найм","hire",C["blue"]),
        ("Порты","ports",C["surf2"]),("Риск","risk",C["red_l"]),
    ]:
        vnames = [r["Судно"] for r in rows]
        y_vals = []
        for r in rows:
            vv = dict(next((v for v in vessels if v["name"]==r["Судно"]),{}))
            vv["heat_mt_day"] = heat_f
            rc = voyage_cost(load, disch, qty_f, int(month or 0),
                             vv, bunker, contract)
            y_vals.append(getattr(rc, "__missing__", None) or
                           (rc[key] if rc else 0))
        fig2.add_trace(go.Bar(name=label, x=vnames, y=y_vals,
            marker_color=col, marker_line_color=C["ink"],marker_line_width=1))
    fig2.update_layout(barmode="stack", height=220,
        margin=dict(l=40,r=10,t=10,b=60),
        paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
        font=dict(family=F,size=10,color=C["ink"]),
        legend=dict(orientation="h",y=-0.35),
        xaxis=dict(showgrid=False,tickangle=-30),
        yaxis=dict(title="$",gridcolor=C["ink_l"],tickprefix="$"))

    return html.Div([
        card([
            html.Div(f"Лучший вариант по $/т: {best}",
                style={"fontWeight":"800","color":C["green"],"marginBottom":"8px",
                       "fontSize":"13px"}),
            html.Div([
                dcc.Graph(figure=fig, config={"displayModeBar":False},
                    style={"flex":"1"}),
                dcc.Graph(figure=fig2, config={"displayModeBar":False},
                    style={"flex":"1"}),
            ], style={"display":"flex","gap":"10px"}),
        ]),
        card([dash_table.DataTable(data=rows,
            columns=[{"name":c,"id":c} for c in rows[0].keys()],
            style_data_conditional=[
                {"if":{"row_index":"odd"},"backgroundColor":C["surf2"]},
                {"if":{"filter_query":'{Судно} = "'+best+'"'},
                 "backgroundColor":C["green_l"],"fontWeight":"700"},
            ],
            **{k:v for k,v in TBL.items() if k != "style_data_conditional"})]),
    ])

# ==========================================================================
#  РАСПИСАНИЕ
# ==========================================================================
@app.callback(Output("schedule-results","children"),
              Input("btn-schedule","n_clicks"),
              State("sch-vessel","value"),
              State("sch-start","value"),
              State("s-calc","data"),
              State("s-vessels","data"),
              prevent_initial_call=True)
def build_schedule(n, vessel_name, start_str, calc, vessels):
    shp_res = (calc or {}).get("shipments",[])
    vs_shps = [r for r in shp_res if r.get("vessel") == vessel_name]
    if not vs_shps:
        return card([html.P("Нет отправок для этого судна. Рассчитайте отправки.",
            style={"color":C["muted"]})])

    try:
        cur = datetime.strptime(start_str.strip(), "%Y-%m-%d")
    except Exception:
        cur = datetime(2026,1,1)

    events = []
    for r in sorted(vs_shps, key=lambda x: (x.get("month") or 0,
                                              x.get("date_start") or "")):
        if r.get("date_start"):
            try:
                cur = datetime.strptime(r["date_start"], "%Y-%m-%d")
            except Exception:
                pass
        end = cur + timedelta(days=r["total_days"])
        events.append({
            "id": r.get("shipment_id",""),
            "load": r["load"], "disch": r["disch"],
            "start": cur.strftime("%Y-%m-%d"),
            "end":   end.strftime("%Y-%m-%d"),
            "days":  r["total_days"],
            "qty":   r.get("qty",0),
            "cpt":   r["cpt"],
            "nm":    r["nm"],
            "total": r["total"],
        })
        cur = end

    # Gantt через Scattergeo‐like → используем go.Figure со scatter
    fig = go.Figure()
    for i, ev in enumerate(events):
        s = datetime.strptime(ev["start"],"%Y-%m-%d")
        e = datetime.strptime(ev["end"],  "%Y-%m-%d")
        color = C["blue_l"] if i%2==0 else C["blue"]
        fig.add_trace(go.Scatter(
            x=[s,e,e,s,s], y=[i+0.1,i+0.1,i+0.9,i+0.9,i+0.1],
            fill="toself", fillcolor=color,
            line=dict(color=C["ink"],width=1),
            mode="lines",
            name=ev["id"],
            hovertemplate=(f"<b>{ev['id']}</b><br>"
                           f"{ev['load']} → {ev['disch']}<br>"
                           f"{ev['start']} – {ev['end']}<br>"
                           f"{ev['days']:.1f} дней · {ev['nm']:,} nm<br>"
                           f"${ev['cpt']:,.2f}/т · ${ev['total']:,.0f}<extra></extra>"),
            showlegend=False,
        ))
        fig.add_annotation(
            x=s+(e-s)/2, y=i+0.5,
            text=f"{ev['id']} · {ev['load'][:8]}→{ev['disch'][:8]} · ${ev['cpt']:.2f}/т",
            showarrow=False, font=dict(size=9,color=C["ink"],family=F),
            xanchor="center",yanchor="middle",
        )

    fig.update_layout(
        height=max(280, 50*len(events)+40),
        margin=dict(l=40,r=10,t=10,b=40),
        paper_bgcolor=C["surface"], plot_bgcolor=C["surf2"],
        font=dict(family=F,size=10,color=C["ink"]),
        xaxis=dict(type="date",showgrid=True,gridcolor=C["ink_l"],
                   tickformat="%b %Y",dtick="M1"),
        yaxis=dict(visible=False,range=[-0.2,len(events)+0.2]),
    )

    total_days = sum(e["days"] for e in events)
    total_cost = sum(e["total"] for e in events)
    total_qty  = sum(e["qty"]   for e in events)

    tbl = dash_table.DataTable(
        data=[{
            "ID": ev["id"], "Маршрут": f"{ev['load']} → {ev['disch']}",
            "Начало": ev["start"], "Конец": ev["end"],
            "Дни": ev["days"], "mt": f"{ev['qty']:,.0f}",
            "$/т": f"${ev['cpt']:,.2f}", "Затраты $": money(ev["total"]),
            "nm": f"{ev['nm']:,}",
        } for ev in events],
        columns=[{"name":c,"id":c} for c in
            ["ID","Маршрут","Начало","Конец","Дни","mt","$/т","Затраты $","nm"]],
        **TBL)

    return html.Div([
        html.Div([
            kpi("Судно",       vessel_name, ""),
            kpi("Рейсов",      str(len(events))),
            kpi("Всего дней",  f"{total_days:.0f}"),
            kpi("Объём mt",    f"{total_qty:,.0f}"),
            kpi("Ср. $/т",     f"${total_cost/total_qty:.2f}" if total_qty else "—"),
        ], style={"display":"flex","gap":"8px","flexWrap":"wrap","marginBottom":"10px"}),
        card([sec(f"Расписание: {vessel_name}"),
              dcc.Graph(figure=fig, config={"displayModeBar":False})]),
        card([tbl]),
    ])

# ==========================================================================
#  КАРТА — обновление региона
# ==========================================================================
@app.callback(Output("main-map","figure"),
              Input("btn-map-refresh","n_clicks"),
              State("map-region","value"),
              State("s-calc","data"),
              prevent_initial_call=True)
def refresh_map(n, region, calc):
    shp_res = (calc or {}).get("shipments",[])
    return build_map(shp_res, region or "Мир целиком")

# ==========================================================================
#  EXCEL ИМПОРТ / ЭКСПОРТ
# ==========================================================================
XL_MAP = {
    "код":"code","code":"code",
    "название":"name","наименование":"name","name":"name",
    "группа":"group","group":"group",
    "плотн.":"density","плотность":"density","density":"density",
    "тепл.°c":"heat_deg","подогрев":"heat_deg","heat_deg":"heat_deg",
    "mt/d":"heat_mt","heat_mt":"heat_mt",
    "мин.партия":"parcel_min","parcel_min":"parcel_min",
    "макс.партия":"parcel_max","parcel_max":"parcel_max",
    "танки":"tanks","tanks":"tanks",
    "заметки":"notes","примечания":"notes","notes":"notes",
    "порт":"port","port":"port",
    "vlsfo":"vlsfo","mgo":"mgo",
    "port dues":"port_dues","portovye sbory":"port_dues","port_dues":"port_dues",
    "canal dues":"canal_dues","canal_dues":"canal_dues",
    "agents":"agents","агенты":"agents",
    "судно":"name","vessel":"name",
    "класс":"cls","cls":"cls",
    "intake mt":"intake","intake":"intake",
    "уз":"speed","скорость":"speed","speed":"speed",
    "vlsfo/d":"cons_sea","cons_sea":"cons_sea",
    "mgo/d":"cons_mgo","cons_mgo":"cons_mgo",
    "порт/d":"cons_port","cons_port":"cons_port",
    "bunk.tank":"tank","tank":"tank",
    "load.h":"load_hours","load_hours":"load_hours",
    "disch.h":"disch_hours","disch_hours":"disch_hours",
    "tc $/d":"tc_rate","tc_rate":"tc_rate",
    "opex $/d":"opex_day","opex_day":"opex_day",
    "стоим.$":"value","value":"value",
    "оффхайр":"offhire_days","offhire_days":"offhire_days",
    "heat mt/d":"heat_mt_day","heat_mt_day":"heat_mt_day",
    "статус":"status","status":"status",
}

def xl_read_sheet(wb, names):
    for nm in names:
        if nm in wb.sheetnames:
            ws = wb[nm]
            rows = list(ws.iter_rows(values_only=True))
            if not rows: return None
            heads = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
            keys  = [XL_MAP.get(h) for h in heads]
            out = []
            for r in rows[1:]:
                if all(v is None for v in r): continue
                rec = {k:v for k,v in zip(keys,r) if k and v is not None}
                if rec: out.append(rec)
            return out
    return None

def parse_upload(contents):
    if not contents: return None
    raw = base64.b64decode(contents.split(",")[1])
    return openpyxl.load_workbook(io.BytesIO(raw), data_only=True)

@app.callback(Output("s-cargoes","data",allow_duplicate=True),
              Input("up-cargoes","contents"), State("s-cargoes","data"),
              prevent_initial_call=True)
def up_cargoes(contents, cur):
    wb = parse_upload(contents)
    if not wb: raise PreventUpdate
    d = xl_read_sheet(wb, ["Грузы","Cargoes"])
    return d if d else cur

@app.callback(Output("s-bunker","data",allow_duplicate=True),
              Input("up-bunker","contents"), State("s-bunker","data"),
              prevent_initial_call=True)
def up_bunker(contents, cur):
    wb = parse_upload(contents)
    if not wb: raise PreventUpdate
    d = xl_read_sheet(wb, ["Бункер","Bunker"])
    return d if d else cur

@app.callback(Output("s-vessels","data",allow_duplicate=True),
              Input("up-vessels","contents"), State("s-vessels","data"),
              prevent_initial_call=True)
def up_vessels(contents, cur):
    wb = parse_upload(contents)
    if not wb: raise PreventUpdate
    d = xl_read_sheet(wb, ["Суда","Vessels"])
    return d if d else cur

@app.callback(Output("dl-template","data"),
              Input("btn-template","n_clicks"),
              prevent_initial_call=True)
def dl_template(n):
    wb = openpyxl.Workbook()
    INK="15202B"; BLU="38648A"; WSH="DFE7EE"
    def hrow(ws, vals):
        ws.append(vals)
        for c in ws[ws.max_row]:
            c.font = Font(bold=True,color="FFFFFF",size=10)
            c.fill = PatternFill("solid",fgColor=BLU)
            c.alignment = Alignment(horizontal="center")
            c.border = Border(left=Side(style="thin",color=INK),
                right=Side(style="thin",color=INK),
                top=Side(style="thin",color=INK),
                bottom=Side(style="thin",color=INK))
    def drow(ws, vals):
        ws.append(vals)
        ri = ws.max_row
        for c in ws[ri]:
            c.font = Font(size=10)
            if ri%2==0: c.fill = PatternFill("solid",fgColor=WSH)
            c.border = Border(left=Side(style="thin",color="C3CFD9"),
                right=Side(style="thin",color="C3CFD9"),
                top=Side(style="thin",color="C3CFD9"),
                bottom=Side(style="thin",color="C3CFD9"))

    ws = wb.active; ws.title="Грузы"
    hrow(ws, ["Код","Название","Группа","Плотн.","Тепл.°C","mt/d",
              "Мин.партия","Макс.партия","Танки","Заметки"])
    for c in DEFAULT_CARGOES:
        drow(ws, [c["code"],c["name"],c["group"],c["density"],
                  c["heat_deg"],c["heat_mt"],c["parcel_min"],
                  c["parcel_max"],c["tanks"],c["notes"]])
    for i in range(1,11): ws.column_dimensions[get_column_letter(i)].width=16

    wb2 = wb.create_sheet("Бункер")
    hrow(wb2, ["Порт","VLSFO","MGO","Port dues","Canal dues","Agents"])
    for b in DEFAULT_BUNKER:
        drow(wb2, [b["port"],b["vlsfo"],b["mgo"],
                   b["port_dues"],b["canal_dues"],b["agents"]])
    for i in range(1,7): wb2.column_dimensions[get_column_letter(i)].width=18

    wb3 = wb.create_sheet("Суда")
    hrow(wb3, ["Судно","Класс","Intake mt","Уз","VLSFO/d","MGO/d","Порт/d",
               "Bunk.tank","Load.h","Disch.h","TC $/d","Opex $/d","Стоим.$",
               "Оффхайр","Heat mt/d","Статус","Заметки"])
    for v in DEFAULT_VESSELS:
        drow(wb3, [v["name"],v["cls"],v["intake"],v["speed"],v["cons_sea"],
                   v["cons_mgo"],v["cons_port"],v["tank"],v["load_hours"],
                   v["disch_hours"],v["tc_rate"],v["opex_day"],v["value"],
                   v["offhire_days"],v["heat_mt_day"],v["status"],v["notes"]])
    for i in range(1,18): wb3.column_dimensions[get_column_letter(i)].width=14

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return dcc.send_bytes(lambda _: buf.getvalue(), "refs_template.xlsx")

# ==========================================================================
#  XLSX ЭКСПОРТ
# ==========================================================================
@app.callback(Output("dl-xlsx","data"),
              Input("btn-dl-xlsx","n_clicks"),
              State("s-calc","data"),
              State("s-vessels","data"), State("s-bunker","data"),
              prevent_initial_call=True)
def dl_xlsx(n, calc, vessels, bunker):
    wb = openpyxl.Workbook()
    INK="15202B"; BLU="38648A"; WSH="DFE7EE"
    def hrow(ws,vals):
        ws.append(vals)
        for c in ws[ws.max_row]:
            c.font=Font(bold=True,color="FFFFFF",size=10)
            c.fill=PatternFill("solid",fgColor=BLU)
            c.alignment=Alignment(horizontal="center")
    def drow(ws,vals):
        ws.append(vals)
        ri=ws.max_row
        for c in ws[ri]:
            c.font=Font(size=10)
            if ri%2==0: c.fill=PatternFill("solid",fgColor=WSH)

    ws = wb.active; ws.title="Отправки"
    hrow(ws,["ID","Маршрут","Судно","Контракт","nm","Дни","mt",
             "Бункер $","Порты $","Найм $","Риск $","Итого $","$/т","SECA","Риск-зоны"])
    for ci,w in zip(range(1,16),[8,18,16,14,8,6,10,10,10,10,10,10,8,20,20]):
        ws.column_dimensions[get_column_letter(ci)].width=w
    shp_res = (calc or {}).get("shipments",[])
    for r in shp_res:
        drow(ws,[r.get("shipment_id",""),f"{r['load']} → {r['disch']}",
                 r.get("vessel",""),r.get("contract",""),r["nm"],
                 r["total_days"],r.get("qty",0),r["bunker"],r["ports"],
                 r["hire"],r["risk"],r["total"],r["cpt"],
                 "; ".join(r.get("seca",[])),"; ".join(r.get("risk_zones",[]))])

    ws2 = wb.create_sheet("Бюджет")
    hrow(ws2,["Месяц","Затраты $","Бункер $","Найм $","Риск $","$/т","Объём mt","Судо-дни"])
    for ci in range(1,9): ws2.column_dimensions[get_column_letter(ci)].width=16
    cost_m=[0.]*12; bunk_m=[0.]*12; hire_m=[0.]*12; risk_m=[0.]*12
    qty_m=[0.]*12; days_m=[0.]*12; cpt_s=[0.]*12; cpt_n=[0]*12
    for r in shp_res:
        m=int(r.get("month") or 0)
        idxs=[m-1] if 1<=m<=12 else list(range(12))
        w=1. if len(idxs)==1 else 1./12
        for i in idxs:
            cost_m[i]+=r["total"]*w; bunk_m[i]+=r["bunker"]*w
            hire_m[i]+=r["hire"]*w;  risk_m[i]+=r["risk"]*w
            qty_m[i]+=r.get("qty",0)*w; days_m[i]+=r["total_days"]*w
            cpt_s[i]+=r["cpt"]*w; cpt_n[i]+=1
    for i in range(12):
        drow(ws2,[MONTH_NAMES[i],round(cost_m[i]),round(bunk_m[i]),
                  round(hire_m[i]),round(risk_m[i]),
                  round(cpt_s[i]/cpt_n[i],2) if cpt_n[i] else 0,
                  round(qty_m[i]),round(days_m[i],1)])
    drow(ws2,["ИТОГО",round(sum(cost_m)),round(sum(bunk_m)),round(sum(hire_m)),
              round(sum(risk_m)),round(sum(cost_m)/sum(qty_m),2) if sum(qty_m) else 0,
              round(sum(qty_m)),round(sum(days_m),1)])
    for ci in range(1,9):
        c=ws2.cell(14,ci); c.font=Font(bold=True,size=10,color="FFFFFF")
        c.fill=PatternFill("solid",fgColor=INK)

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return dcc.send_bytes(lambda _: buf.getvalue(), "edible_oils_costs.xlsx")

@app.callback(Output("dl-json","data"),
              Input("btn-dl-json","n_clicks"),
              State("s-calc","data"), State("s-shipments","data"),
              State("s-vessels","data"), State("s-bunker","data"),
              State("s-cargoes","data"), State("s-routes","data"),
              prevent_initial_call=True)
def dl_json_cb(n, calc, shp, vessels, bunker, cargoes, routes):
    payload={
        "meta":{"app":"Edible Oils Charterer Planner","perspective":"charterer",
                "generated":datetime.utcnow().isoformat()+"Z"},
        "cargoes":cargoes or DEFAULT_CARGOES,
        "vessels":vessels or DEFAULT_VESSELS,
        "bunker_prices":bunker or DEFAULT_BUNKER,
        "routes":routes or DEFAULT_ROUTES,
        "shipments":shp or [],
        "calc_results":(calc or {}).get("shipments",[]),
    }
    blob=json.dumps(payload,ensure_ascii=False,indent=2).encode()
    return dcc.send_bytes(lambda _: blob,"edible_oils_dataset.json")

# ==========================================================================
#  СЦЕНАРИИ
# ==========================================================================
@app.callback(Output("s-scenarios","data",allow_duplicate=True),
              Input("btn-save-sc","n_clicks"),
              State("sc-name","value"),
              State("s-shipments","data"), State("s-vessels","data"),
              State("s-bunker","data"), State("s-cargoes","data"),
              State("s-routes","data"), State("s-calc","data"),
              State("s-scenarios","data"),
              prevent_initial_call=True)
def save_sc(n, name, shp, vessels, bunker, cargoes, routes, calc, scenarios):
    nm = (name or "").strip() or f"Сценарий {len(scenarios or [])+1}"
    snap={"name":nm,"saved":datetime.now().strftime("%Y-%m-%d %H:%M"),
          "shipments":shp or [],"vessels":vessels or DEFAULT_VESSELS,
          "bunker":bunker or DEFAULT_BUNKER,"cargoes":cargoes or DEFAULT_CARGOES,
          "routes":routes or DEFAULT_ROUTES,
          "calc":(calc or {}).get("shipments",[])}
    out=[s for s in (scenarios or []) if s.get("name")!=nm]
    out.append(snap)
    return out

@app.callback(Output("s-scenarios","data",allow_duplicate=True),
              Input({"type":"sc-del","index":ALL},"n_clicks"),
              State("s-scenarios","data"), prevent_initial_call=True)
def del_sc(n, scenarios):
    trig = ctx.triggered_id
    if not trig: raise PreventUpdate
    out = list(scenarios or [])
    idx = trig["index"]
    if 0<=idx<len(out): out.pop(idx)
    return out

@app.callback(Output("s-snap","data"),
              Input({"type":"sc-load","index":ALL},"n_clicks"),
              State("s-scenarios","data"), prevent_initial_call=True)
def pick_sc(n, scenarios):
    trig = ctx.triggered_id
    if not trig: raise PreventUpdate
    idx = trig["index"]
    if scenarios and 0<=idx<len(scenarios): return scenarios[idx]
    raise PreventUpdate

@app.callback(
    Output("s-cargoes","data",  allow_duplicate=True),
    Output("s-bunker","data",   allow_duplicate=True),
    Output("s-vessels","data",  allow_duplicate=True),
    Output("s-shipments","data",allow_duplicate=True),
    Output("s-routes","data",   allow_duplicate=True),
    Output("s-calc","data",     allow_duplicate=True),
    Input("s-snap","data"), prevent_initial_call=True)
def apply_sc(snap):
    if not snap: raise PreventUpdate
    results, _ = run_calc(snap.get("shipments",[]),
        snap.get("vessels",DEFAULT_VESSELS),
        snap.get("bunker",DEFAULT_BUNKER))
    return (snap.get("cargoes",DEFAULT_CARGOES),
            snap.get("bunker",DEFAULT_BUNKER),
            snap.get("vessels",DEFAULT_VESSELS),
            snap.get("shipments",[]),
            snap.get("routes",DEFAULT_ROUTES),
            {"shipments":results})

@app.callback(Output("dl-scenario","data"),
              Input({"type":"sc-exp","index":ALL},"n_clicks"),
              State("s-scenarios","data"), prevent_initial_call=True)
def exp_sc(n, scenarios):
    trig = ctx.triggered_id
    if not trig: raise PreventUpdate
    idx = trig["index"]
    if not scenarios or idx>=len(scenarios): raise PreventUpdate
    snap = scenarios[idx]
    blob = json.dumps(snap,ensure_ascii=False,indent=2).encode()
    safe = re.sub(r"[^A-Za-z0-9_.-]+","_",snap.get("name","sc"))
    return dcc.send_bytes(lambda _: blob, f"scenario_{safe}.json")

@app.callback(Output("dl-scenarios-all","data"),
              Input("btn-exp-all","n_clicks"),
              State("s-scenarios","data"), prevent_initial_call=True)
def exp_all_sc(n, scenarios):
    blob=json.dumps({"scenarios":scenarios or []},ensure_ascii=False,indent=2).encode()
    return dcc.send_bytes(lambda _: blob,"scenarios_all.json")

@app.callback(Output("s-scenarios","data",allow_duplicate=True),
              Input("up-scenario","contents"),
              State("s-scenarios","data"), prevent_initial_call=True)
def imp_sc(contents, scenarios):
    if not contents: raise PreventUpdate
    raw = base64.b64decode(contents.split(",")[1])
    data = json.loads(raw.decode("utf-8"))
    if isinstance(data,dict) and "scenarios" in data:
        incoming = data["scenarios"]
    elif isinstance(data,dict):
        incoming = [data]
    else:
        incoming = data if isinstance(data,list) else []
    out = list(scenarios or [])
    names = {s.get("name") for s in out}
    for s in incoming:
        if not isinstance(s,dict): continue
        if s.get("name") in names:
            s["name"] = str(s.get("name","sc"))+" (импорт)"
        out.append(s)
    return out

# ==========================================================================
#  ЗАПУСК
# ==========================================================================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    app.run(host="0.0.0.0", port=port, debug=False,
            use_reloader=False, dev_tools_hot_reload=False)
