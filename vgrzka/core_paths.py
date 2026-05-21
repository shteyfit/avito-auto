from pathlib import Path

# ========== БАЗОВЫЕ ДИРЕКТОРИИ ==========
# Папка, где лежит программа (prgrm)
BASE_DIR = Path(__file__).resolve().parent

# DIR_DIR: сырые входные файлы от тебя и поставщиков
#   - свежая выгрузка остатков из 1С (например "26.10.25.xlsx")
#   - каталог поставщика Kolobox (catalog_tyres.xlsx)
#   - полный справочник шин Kolobox (tyres.xlsx)
#   - выгрузка 4tochki (Forto4ki_tires_detailed.xlsx)
DIR_DIR = BASE_DIR / "dir"

# SOP_DIR: рабочие и подготовленные таблицы
#   - result_with_articul.xlsx
#   - kolobox_prepared.xlsx
#   - 4tochki_prepared.xlsx
#   - товарынаАвито.xlsx
#   - updated_avito.xlsx
SOP_DIR = BASE_DIR / "sopostav"


def ensure_directories():
    """Создаём каталоги, если их нет."""
    DIR_DIR.mkdir(parents=True, exist_ok=True)
    SOP_DIR.mkdir(parents=True, exist_ok=True)


# Совместимость со старыми именами
RAW_DIR = DIR_DIR
SUP_DIR = DIR_DIR          # старый код мог звать SUP_DIR -> это то же, что DIR_DIR
PREPARED_DIR = SOP_DIR     # иногда старый код называл sopostav 'PREPARED_DIR'
BASE_PATH = BASE_DIR       # на всякий случай


# ========== СЫРЫЕ ФАЙЛЫ (входные) ==========

# 1. Kolobox сырой каталог (актуальные остатки поставщика)
KOLOBOX_RAW_CATALOG = DIR_DIR / "catalog_tyres.xlsx"

# 2. Полная НОМЕНКЛАТУРА Kolobox (большой справочник всех шин)
#    В GUI это получается кнопкой "Получить номенклатуру Kolobox".
#    Этот файл нам нужен как справочник параметров для новых товаров.
TYRES_REFERENCE_FILE = DIR_DIR / "tyres.xlsx"

# 3. Сырые остатки 4tochki (детализированный файл после API)
FOURTOCHKI_RAW = DIR_DIR / "Forto4ki_tires_detailed.xlsx"

# 4. NashArticul.xlsx — твоя база "все товары, когда-либо продаваемые у нас"
#    обычно ты выбираешь её вручную, но пусть путь по умолчанию будет тут:
NASH_ARTICUL_FILE = SOP_DIR / "NashArticul.xlsx"


# ========== ПОДГОТОВЛЕННЫЕ ФАЙЛЫ ПОСТАВЩИКОВ ==========

# kolobox_prepared.xlsx — результат нормализации Kolobox
# c расчетом ЦенаАвито = закупка*1.1 и округлением вверх до десятка
KOLOBOX_PREPARED = SOP_DIR / "kolobox_prepared.xlsx"

# 4tochki_prepared.xlsx — результат нормализации 4tochki
FOURTOCHKI_PREPARED = SOP_DIR / "4tochki_prepared.xlsx"

# совместимость со старыми именами:
KOLOBOX_PREPARED_FILE = KOLOBOX_PREPARED
KOLOBOX_FILE_PREPARED = KOLOBOX_PREPARED
FOURTOCHKI_PREPARED_FILE = FOURTOCHKI_PREPARED
FOURTOCHKI_FILE_PREPARED = FOURTOCHKI_PREPARED
FOURTOchKI_FILE_PREPARED = FOURTOCHKI_PREPARED  # опечатка в старом коде
T4_PREPARED_FILE = FOURTOCHKI_PREPARED          # старое имя для 4точки


# ========== НАШИ ОСТАТКИ (после вкладки "Сбор артикулов") ==========

# Текущий файл остатков с артикулами:
#   Id | Название товара | Статус | Цена | Количество
RESULT_WITH_ARTICUL = SOP_DIR / "result_with_articul.xlsx"

# Предыдущая версия (старые цены, вручную выставленные):
RESULT_WITH_ARTICUL_PREV = SOP_DIR / "result_with_articul_prev.xlsx"

# совместимость со старыми именами:
OUR_STOCK_FILE = RESULT_WITH_ARTICUL
OUR_STOCK_FILE_PREV = RESULT_WITH_ARTICUL_PREV
PREVIOUS_RESULT_FILE = RESULT_WITH_ARTICUL_PREV


# ========== ЦЕНЫ ==========

# Файл с новыми ценами (ручная переоценка):
#   Название | Цена
NEW_PRICE_FILE = SOP_DIR / "newprice.xlsx"
NEW_PRICE_LIST_FILE = NEW_PRICE_FILE  # алиас для старых вызовов


# ========== ВЫГРУЗКА АВИТО ==========

# Исходный файл с позициями, которые УЖЕ на Авито
# (у них заполнены бренд/модель/описание/категория и т.д.)
AVITO_BASE_FILE = SOP_DIR / "товарынаАвито.xlsx"

# Финальная таблица, готовая к заливке на Авито после сопоставления
UPDATED_AVITO_FILE = SOP_DIR / "updated_avito.xlsx"

# совместимость со старыми именами:
AVITO_FILE = AVITO_BASE_FILE
OUTPUT_UPDATED_AVITO = UPDATED_AVITO_FILE
OUTPUT_UPDATED_AVITO_FILE = UPDATED_AVITO_FILE
UPDATED_AVITO = UPDATED_AVITO_FILE
