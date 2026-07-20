"""
Конфигурация автоматизации отчётов закупа для QuickBooks.

Все справочники и константы вынесены сюда для удобства поддержки
без правки кода движка.
"""

# ---------------------------------------------------------------------------
# Маппинг площадок: имя → фильтры в сырых отчётах
# ---------------------------------------------------------------------------
# r1, r2, genba — могут быть строкой ИЛИ списком строк (для комбинированных площадок)
# Имя синтетической колонки с вычисленной площадкой
# (добавляется к R1/R2 в pipeline после загрузки; ручную 'площадка' игнорируем — она была вручную проставляемой меткой,
# а с апреля 2026 в выгрузке вообще исчезла)
SYNTH_PLOSHADKA = "_ploshadka"

# ---------------------------------------------------------------------------
# Правила вычисления площадки по колонкам Партнёр + Ключ куплен в сток + Поставщик
# ---------------------------------------------------------------------------
# Партнёр → (zone_zakup, zone_peremeshchenie).
# Если 'Ключ куплен в сток' = ДА → берётся zone_peremeshchenie, иначе zone_zakup.
#
# Cost*-партнёры — это технические счета для учёта стоковых расходов, эквивалентные
# Stock*-партнёрам (см. v13). По апрельской разметке: CostGB ↔ StockGB, CostChinaplay ↔ StockChinaPlay.
PARTNER_TO_ZONE = {
    "MP_Plati":           ("закуп плати",   "перемещение на плати"),
    "MP_Kinguin":         ("закуп кингвин", "перемещение кингвин"),
    "MP_Eneba":           ("закуп энеба",   "перемещение енеба"),
    "MP_G2A":             ("закуп г2а",     "перемещение на г2а"),
    "MP_Driffle":         ("закуп дриффл",  "перемещение дриффл"),
    "ChinaSteamPY":       ("закуп тао",     "перемещение на тао"),
    "ChinaPlayTaoBao":    ("закуп тао",     "перемещение на тао"),
    "StockB2B":           ("закуп b2b",     "закуп b2b"),
    "CostB2B":            ("закуп b2b",     "cost b2b"),       # появился в апреле
    "StockGB":            ("закуп гб",      "перемещение на гб"),
    "CostGB":             ("закуп гб",      "перемещение на гб"),
    "StockChinaPlay":     ("закуп чайна",   "закуп чайна"),
    "CostChinaplay":      ("закуп чайна",   "costchinaplay"),
    "Gamersbase_WW":      ("продажи гб",    "продажи гб"),
}

# Партнёр Rokky Platform / (CN) yueshangshuma может попадать в три зоны
# в зависимости от поставщика:
#   - содержит '(GB_Stock)' или 'Enaza-Games' (RUB) → продажи гб
#   - валюта CNY и поставщик стоковый/CNY-Team17 → продажи на чайнаплей
#   - всё прочее → Продажи б2б
DUAL_USE_PARTNERS = {"Rokky Platform", "(CN) yueshangshuma"}
GB_SUPPLIER_MARKERS = ["(GB_Stock)", "GamersBase", "Enaza-Games"]
CHINAPLAY_SUPPLIER_MARKERS = ["(Stock)", "Team17"]
DEFAULT_B2B_ZONE = "Продажи б2б"

# ---------------------------------------------------------------------------
# Маппинг площадок: имя → фильтры в сырых отчётах
# ---------------------------------------------------------------------------
# r1, r2, genba — могут быть строкой ИЛИ списком строк (для комбинированных площадок)
PLOSHADKA_MAP = {
    "Plati":      {"r1": "закуп плати",   "r2": "закуп плати",                    "genba": "плати"},
    "Kinguin":    {"r1": "закуп кингвин", "r2": "закуп кингвин",                  "genba": "кингвин"},
    "Eneba":      {"r1": "закуп энеба",   "r2": "закуп энеба",                    "genba": "eneba"},
    # G2A: в марте было 'закуп г2а' (русские буквы), с апреля — 'Закуп G2A' (английские G2A).
    # Поддерживаем оба варианта, чтобы исторические выгрузки тоже работали.
    "G2A":        {"r1": ["закуп г2а", "закуп g2a"], "r2": ["закуп г2а", "закуп g2a"], "genba": "g2a"},
    # Driffle: до июня 2026 закуп шёл только в R2, поэтому r1=None. С июня в R1
    # появились стоковые закупы MP_Driffle (in_stock=НЕТ). R1 и R2 непересекающиеся
    # (∩ по номерам заказов = 0), поэтому чтение R1 не задваивает.
    "Driffle":    {"r1": "закуп дриффл",   "r2": "закуп дриффл",                   "genba": "driffle"},
    "Tao":        {"r1": "закуп тао",     "r2": "закуп тао",                      "genba": "тао"},
    "ChinaPlay":  {"r1": ["закуп чайна", "costchinaplay"], "r2": ["закуп чайна", "costchinaplay"], "genba": ["chinaplay", "costchinaplay"]},
    "B2B":        {"r1": "продажи б2б",   "r2": ["закуп b2b", "Продажи б2б"],     "genba": "b2b",
                   # Закуп PLAION учитывается весь, даже если ключи передали на другие площадки:
                   # эталон агрегирует PLAION по всем зонам, не только по продажам б2б.
                   "extra_supplier_substrings": ["PLAION"]},
    "GamersBase": {"r1": ["закуп гб", "costgb"], "r2": ["закуп гб", "costgb"],     "genba": ["gb", "costgb"]},
}

# ---------------------------------------------------------------------------
# Сопоставление сырых имён поставщиков → группа в финальном своде
# ---------------------------------------------------------------------------
SUPPLIER_MAPPING = {
    # Стандартные издатели
    "Hooded Horse":        "Hooded Horse",
    "Nacon (Point Nexus)": "Nacon",
    "Nacon":               "Nacon",
    "Team17":              "Team17",
    "Team 17":             "Team17",  # вариант в Tao
    "Owlcat Games":        "Owlcat Games",
    "Green Man Gaming":    "Green Man Gaming",
    "ALAWAR":              "ALAWAR",
    "Fulqrum Publishing":  "Fulqrum Publishing",
    "Offworld Industries": "Offworld Industries",
    "THQ Nordic Games":    "THQ Nordic Games",
    "Stunlock Studios":    "Stunlock Studios",
    "Stunlock Studios AB": "Stunlock Studios",  # вариант в Tao
    "DOOR 407":            "DOOR 407",
    "Iceberg Interactive": "Iceberg Interactive",
    "MINTROCKET":          "MINTROCKET",
    "Aspyr":               "Aspyr",
    "Shiravune":           "Shiravune",
    "ArtDock":             "ArtDock",
    "Gamersky":            "Gamersky",
    "Gamersky Games":      "Gamersky",
    "Gamersky games":      "Gamersky",
    # Особые имена групп
    "Ytopia":              "YTOPIA LLC",
    "YTOPIA":              "YTOPIA LLC",
    # CNY-поставщики (см. CNY_SUPPLIERS ниже)
    "Kishmish Games":      "Kishmish Games",
    "One More Time":       "One More Time",
    "Callback Games":      "Callback Games",
    # Новые из Tao/ChinaPlay
    "Daedalic":                    "Daedalic",
    "DAEDALIC ENTERTAINMENT GMBH": "Daedalic",
    "META Publishing":             "META Publishing",
    "Quantic Dream":               "Quantic Dream",
    "Quantic Dream (Point Nexus)": "Quantic Dream",
    "QUANTIC DREAM":               "Quantic Dream",
    "Thunderful Publishing":       "Thunderful Publishing",
    "MY.GAMES":                    "MY.GAMES",
    "Top Hat Studios":             "Top Hat Studios",
    # B2B-поставщики (исторически шли отдельными инвойсами, теперь в биллинге)
    "KRM Teknoloji":               "КRM",  # имя в R1 'продажи б2б' (русская К — как в эталоне)
    # Новые поставщики, появившиеся в апреле 2026
    "Fireshine Games":             "Fireshine Games",
    "FOR-GAMES CR LTD":            "FOR-GAMES CR LTD",
    "Polden Publishing":           "Polden Publishing",
    "CURVE GAMES":                 "CURVE GAMES",
    "Techland":                    "Techland",
    "Strategy First":              "Strategy First",
    "Frontier Developments":       "Frontier Developments",
    "Incenti":                     "Incenti",
    # Неклассифицированные поставщики, найденные в мае 2026 (сверка выхода).
    # Часть — в существующие группы, часть — само-именованные (можно переименовать).
    "STRATEGY FIRST":              "Strategy First",   # верхний регистр из R2
    "Giftcard Pro":                "Giftcard pro LTD",  # без скобок
    "Moogold":                     "Moogold",
    "GamersBase":                  "GamersBase",
    "GFAGAMES":                    "GFAGAMES",
    "MPG":                         "MPG",
    "Cyber Temple Games LLC":      "Cyber Temple Games LLC",
    "3 GAMING PILLARS":            "3 Gaming Pillars",
    "Storytaco Game":              "Storytaco Game",
    "Wired Productions":           "Wired Productions",
}

# Спецсопоставления по подстрокам (применяются раньше префиксного парсинга)
SUPPLIER_SUBSTRING_RULES = [
    # порядок имеет значение: более специфичные правила выше
    ("(Genba)", "Genba"),
    ("Plug-in-Digital", "Plug-in-Digital"),
    ("(PID)",  "Plug-in-Digital"),
    ("(Epay)", "PLN ИГРЫ. (Epay)"),
    # B2B: PLAION приходит как 'EUR GAMES. PLAION (Tier 1/2/3)' — Tier игнорируем
    ("PLAION", "PLAION"),
    # B2B: KRM Teknoloji в R2 — 'TRY/USD GAMES. PlayStation TR (KRM Teknoloji)' и т.п.
    # (русская К — как в эталоне 'КRM')
    ("(KRM Teknoloji)", "КRM"),
    # B2B: Giftcard Pro — 'USD GAMES. Blizzard (Giftcard Pro)'
    ("(Giftcard Pro)", "Giftcard pro LTD"),
    # B2B: Capcom / Embark Studios как Genba-сток
    # (закупка идёт через Genba, в биллинге появляется под брендом продукта)
    ("Capcom (Stock)",        "Genba"),
    ("Embark Studios (Stock)", "Genba"),
    # Embark Studios Tier 1 (Stock) — отдельное имя для USD-Tier 1
    ("Embark Studios Tier",   "Genba"),
    # Найдено в мае 2026: агрегаторы-суффиксы и сток-бренды
    ("(Incenti)", "Incenti"),                    # XXX GAMES. PlayStation/Xbox/… (Incenti)
    ("(VaultN)",  "VaultN"),                     # XXX GAMES. ByteRockers' Games (VaultN)
    ("Fireshine Games", "Fireshine Games"),      # Fireshine Games (Stock)
    ("Genba (Stock)", "Genba"),
    ("Moogold", "Moogold"),                      # Moogold и '… PUBG Mobile (Moogold)'
    ("GamersBase", "GamersBase"),                # GamersBase и 'GamersBase (Stock)'
]

# Точные совпадения (если имя поставщика — это просто слово без префикса)
SUPPLIER_EXACT_RULES = {
    "Genba": "Genba",
    "Epay":  "PLN ИГРЫ. (Epay)",
}

# ---------------------------------------------------------------------------
# CNY-поставщики: цена = (RUB-сумма из биллинга) / RUB_CNY_RATE
# ---------------------------------------------------------------------------
CNY_SUPPLIERS = {"Kishmish Games", "One More Time", "Callback Games"}
RUB_CNY_RATE = 11

# ---------------------------------------------------------------------------
# Валюты в финальном своде по поставщикам
# ---------------------------------------------------------------------------
SUPPLIER_CURRENCY = {
    "Nacon":           "EUR",
    "Daedalic":        "EUR",
    "Quantic Dream":   "EUR",
    "Kishmish Games":  "CNY",
    "One More Time":   "CNY",
    "Callback Games":  "CNY",
    # B2B (по эталону)
    "PLAION":           "EUR",
    "КRM":              "USD",  # источник в TRY, конвертируется в USD
    "Giftcard pro LTD": "USD",
}
DEFAULT_CURRENCY = "USD"

# ---------------------------------------------------------------------------
# FX-фолбэк: средний курс из R2 по валютам.
# Используется ТОЛЬКО когда у позиции пуст 'Курс фиксации в валюте базового поставщика'
# (так бывает для зоны 'Продажи б2б'). Заполняется движком из загруженного R2.
# Структура: {"TRY": 0.0226, ...} — множитель валюта→USD (1 ед. валюты = X USD).
# ---------------------------------------------------------------------------
FX_FALLBACK_DEFAULT = {
    # запасные значения на случай, если в R2 не нашлось ни одной строки с курсом
    "TRY": 0.0228,  # ≈ март 2026 (USD/TRY ≈ 43.9)
}

# ---------------------------------------------------------------------------
# Имена колонок в сырых отчётах
# ---------------------------------------------------------------------------
# Universal Report (R1). До апреля 2026 здесь была ручная колонка 'площадка ' (с пробелом),
# с апреля её нет. Площадка теперь ВЫЧИСЛЯЕТСЯ из 'Партнер' + 'Ключ куплен в сток'.
COLS_R1 = {
    "supplier":    "Поставщик",
    "pid":         "Id продукта (Billing)",
    "prod_name":   "Продукт",
    "base_amount": "Закуп в валюте взаиморасчетов с ПО",
    "base_ccy":    "Валюта базового ПО",
    "prod_amount": "Цена закупа в валюте продукта",
    "prod_ccy":    "Валюта продукта",
    "partner":     "Партнер",                # обязательно для вычисления площадки
    "in_stock":    "Ключ куплен в сток",     # появилась с апреля 2026; до — отсутствует
}

# Universal Report shipped (R2). Аналогично — 'площадка' вычисляется.
COLS_R2 = {
    "supplier":    "Поставщик",
    "pid":         "ID продукта",
    "prod_name":   "Продукт",
    "qty":         "Количество",
    "base_amount": "Сумма в валюте базового поставщика",
    "base_ccy":    "Валюта базового поставщика",
    "prod_amount": "Себестоимость позиции заказа",
    "prod_ccy":    "Валюта покупки у поставщика",
    "fx_rate":     "Курс фиксации в валюте базового поставщика",
    "partner":     "Партнёр",
    "in_stock":    "Ключ куплен в сток",
}

# genbaFile
COLS_GENBA = {
    "ploshadka":   "площадка",
    "pid":         "ID продукта",
    "qty":         "Activation Qty",
    "grand_total": "Grand Total",
}
