import json
import os
import shutil
import pandas as pd
import requests
from bs4 import BeautifulSoup
import reverse_geocoder as rg

# Налаштування
OUTPUT_JSON_FILE = "all_shops.json"
OUTPUT_EXCEL_FILE = "all_shops.xlsx"
BACKUP_JSON_FILE = "all_shops_bcp.json"  # Чистий бекап lifecell

# Файли для зберігання поточних бекапів перед запуском
BACKUP_RUN_JSON = "all_shops_backup_last.json"
BACKUP_RUN_EXCEL = "all_shops_backup_last.xlsx"

# Мапінг англійських назв областей від reverse_geocoder в українські
REGION_MAP = {
    "Cherkasy": "Черкаська",
    "Chernihiv": "Чернігівська",
    "Chernivtsi": "Чернівецька",
    "Dnipropetrovsk": "Дніпропетровська",
    "Dnipro": "Дніпропетровська",
    "Donetsk": "Донецька",
    "Ivano-Frankivsk": "Івано-Франківська",
    "Kharkiv": "Харківська",
    "Kherson": "Херсонська",
    "Khmelnytskyi": "Хмельницька",
    "Kirovohrad": "Кіровоградська",
    "Kyiv": "Київська",
    "Kyiv City": "м. Київ",
    "Luhansk": "Луганська",
    "Lviv": "Львівська",
    "Mykolaiv": "Миколаївська",
    "Odesa": "Одеська",
    "Poltava": "Полтавська",
    "Rivne": "Рівненська",
    "Sumy": "Сумська",
    "Ternopil": "Тернопільська",
    "Vinnytsia": "Вінницька",
    "Volyn": "Волинська",
    "Zakarpatska": "Закарпатська",
    "Zaporizhzhya": "Запорізька",
    "Zhytomyr": "Житомирська",
    "Crimea": "АР Крим"
}

# ------------------------------------------------------------
# 0. МЕХАНІЗМ БЕКАПУ ТА ВІДНОВЛЕННЯ
# ------------------------------------------------------------
def make_initial_backup():
    """Створює резервну копію існуючих файлів перед запуском"""
    print("[БЕКАП] Перевіряємо та створюємо бекап поточних файлів...")
    if os.path.exists(OUTPUT_JSON_FILE):
        try:
            shutil.copy(OUTPUT_JSON_FILE, BACKUP_RUN_JSON)
            print(f" -> Збережено бекап JSON: {BACKUP_RUN_JSON}")
        except Exception as e:
            print(f" -> [ПОМИЛКА] Не вдалося створити бекап JSON: {e}")

    if os.path.exists(OUTPUT_EXCEL_FILE):
        try:
            shutil.copy(OUTPUT_EXCEL_FILE, BACKUP_RUN_EXCEL)
            print(f" -> Збережено бекап Excel: {BACKUP_RUN_EXCEL}")
        except Exception as e:
            print(f" -> [ПОМИЛКА] Не вдалося створити бекап Excel: {e}")

def load_cached_provider_shops(provider_name):
    """
    Якщо живий запит не вдався, шукаємо дані оператора в існуючих бекапах:
    1. Спочатку у щойно створеному BACKUP_RUN_JSON
    2. Потім у файлі OUTPUT_JSON_FILE
    3. Потім у локальному BACKUP_JSON_FILE
    """
    print(f" -> [ФОЛЛБЕК] Спроба зчитати дані для '{provider_name}' з попередніх бекапів...")
    
    for file_path in [BACKUP_RUN_JSON, OUTPUT_JSON_FILE, BACKUP_JSON_FILE]:
        if os.path.exists(file_path):
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    
                    shops = [
                        item for item in data 
                        if isinstance(item, dict) and item.get("provider") == provider_name
                    ]
                    
                    if shops:
                        print(f" -> [ВІДНОВЛЕНО] Знайдено {len(shops)} точок '{provider_name}' у файлі {file_path}")
                        return shops
            except Exception as e:
                print(f" -> Помилка читання файлу {file_path}: {e}")
                
    print(f" -> [УВАГА] Не вдалося знайти бекап для оператора '{provider_name}'. Повертаємо порожній список.")
    return []

# ------------------------------------------------------------
# 1. LIFECELL (Суворо з локального бекап-файлу)
# ------------------------------------------------------------
def fetch_lifecell():
    print(f"\n[1/3] Читаємо магазини lifecell з файлу: {BACKUP_JSON_FILE}...")
    
    if os.path.exists(BACKUP_JSON_FILE):
        try:
            with open(BACKUP_JSON_FILE, "r", encoding="utf-8") as f:
                backup_data = json.load(f)
                
                lifecell_shops = [
                    item for item in backup_data 
                    if isinstance(item, dict) and item.get("provider") == "lifecell"
                ] or backup_data

                print(f" -> Знайдено {len(lifecell_shops)} магазинів lifecell у бекапі.")
                return lifecell_shops
        except Exception as e:
            print(f" -> [ПОМИЛКА] Не вдалося зчитати {BACKUP_JSON_FILE}: {e}")
    else:
        print(f" -> [ПОМИЛКА] Файл {BACKUP_JSON_FILE} відсутній у репозиторії!")

    return load_cached_provider_shops("lifecell")

# ------------------------------------------------------------
# 2. VODAFONE (Живий запит + Фоллбек на бекап)
# ------------------------------------------------------------
def fetch_vodafone():
    print("\n[2/3] Завантажуємо магазини Vodafone з API...")
    url = "https://www.vodafone.ua/shops/kyiv"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "uk-UA,uk;q=0.9,en-US;q=0.8,en;q=0.7"
    }

    try:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")
        script_tag = soup.find("script", id="ng-state")

        if not script_tag or not script_tag.string:
            raise ValueError("ng-state не знайдено або він порожній.")

        data = json.loads(script_tag.string)
        normalized = []

        for main_key, main_value in data.items():
            if isinstance(main_value, dict) and "b" in main_value:
                b_value = main_value["b"]
                if isinstance(b_value, dict) and "data" in b_value:
                    shops = b_value["data"]
                    if isinstance(shops, list):
                        for shop in shops:
                            if isinstance(shop, dict) and "address" in shop and "shopTypes" in shop:
                                lat = shop.get("latitude")
                                lng = shop.get("longitude")

                                if lat is None or lng is None:
                                    continue
                                try:
                                    lat = float(lat)
                                    lng = float(lng)
                                except (TypeError, ValueError):
                                    continue

                                region_city = shop.get("regionCity", {})
                                city_name = region_city.get("name", "") if isinstance(region_city, dict) else ""
                                region_obj = region_city.get("region", {}) if isinstance(region_city, dict) else {}
                                region_name = region_obj.get("name", "") if isinstance(region_obj, dict) else ""

                                working_hours = shop.get("workingHours", "")
                                if working_hours:
                                    working_hours = working_hours.replace("<br>", "\n")

                                normalized.append({
                                    "provider": "vodafone",
                                    "id": shop.get("id"),
                                    "name": shop.get("name") or shop.get("location") or "Vodafone",
                                    "address": shop.get("address", ""),
                                    "city": city_name,
                                    "region": region_name,
                                    "lat": lat,
                                    "lng": lng,
                                    "working_hours": working_hours
                                })

        unique = []
        seen = set()
        for shop in normalized:
            key = (shop["provider"], shop["lat"], shop["lng"], shop["address"])
            if key not in seen:
                seen.add(key)
                unique.append(shop)

        if not unique:
            raise ValueError("Отримано 0 записів після обробки.")

        print(f" -> УСПІХ: Знайдено {len(unique)} магазинів Vodafone.")
        return unique

    except Exception as e:
        print(f" -> [ПОМИЛКА Vodafone]: {e}")
        return load_cached_provider_shops("vodafone")

# ------------------------------------------------------------
# 3. KYIVSTAR (Живий запит + Фоллбек на бекап)
# ------------------------------------------------------------
def fetch_kyivstar():
    print("\n[3/3] Завантажуємо магазини Kyivstar з API...")
    url = "https://kyivstar.ua/api/pos-shops?locale=uk&start=0&limit=10000"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "uk-UA,uk;q=0.9,en;q=0.8",
        "Referer": "https://kyivstar.ua/shops"
    }

    try:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()

        raw = data.get("data", []) if isinstance(data, dict) else data
        normalized = []

        for item in raw:
            if not isinstance(item, dict):
                continue

            pos = item.get("position", {}) or {}
            lat = pos.get("latitude") or item.get("latitude") or item.get("lat")
            lng = pos.get("longitude") or item.get("longitude") or item.get("lng")

            if lat is None or lng is None:
                continue

            try:
                lat = float(lat)
                lng = float(lng)
            except (TypeError, ValueError):
                continue

            street_type = item.get("streetType", "") or ""
            street_name = item.get("streetName", "") or ""
            building = item.get("buildingNumber", "") or ""
            full_address = f"{street_type} {street_name} {building}".strip()

            city_obj = item.get("city", {}) or {}
            city_name = city_obj.get("city", "") if isinstance(city_obj, dict) else ""

            wh_list = item.get("workingHours", [])
            wh_str_list = []
            if isinstance(wh_list, list):
                for wh in wh_list:
                    if isinstance(wh, dict):
                        day = wh.get("dayOfWeek", "")
                        start = (wh.get("startTime") or "")[:5]
                        end = (wh.get("endTime") or "")[:5]
                        is_holiday = wh.get("isHoliday", False)
                        if is_holiday:
                            wh_str_list.append(f"{day}: Вихідний")
                        elif start and end:
                            wh_str_list.append(f"{day}: {start}-{end}")
            
            working_hours = "\n".join(wh_str_list)

            normalized.append({
                "provider": "kyivstar",
                "id": item.get("id"),
                "name": "Kyivstar",
                "address": full_address,
                "city": city_name,
                "region": "",
                "lat": lat,
                "lng": lng,
                "working_hours": working_hours
            })

        if not normalized:
            raise ValueError("Отримано 0 записів від Kyivstar API.")

        print(f" -> УСПІХ: Знайдено {len(normalized)} магазинів Kyivstar.")
        return normalized

    except Exception as e:
        print(f" -> [ПОМИЛКА Kyivstar]: {e}")
        return load_cached_provider_shops("kyivstar")

# ------------------------------------------------------------
# 4. REVERSE GEOCODING (Область за координатами)
# ------------------------------------------------------------
def enrich_with_regions(shops):
    print("\n[Геолокація] Визначаємо області за координатами через reverse_geocoder...")
    
    coords = [(shop["lat"], shop["lng"]) for shop in shops if "lat" in shop and "lng" in shop]
    if not coords:
        return shops

    try:
        results = rg.search(coords)

        for i, shop in enumerate(shops):
            res = results[i]
            admin1 = res.get("admin1", "").replace(" Oblast", "").strip()
            ua_region = REGION_MAP.get(admin1, admin1)
            
            shop["geo_region"] = ua_region
            if not shop.get("region"):
                shop["region"] = ua_region

        print(" -> Області успішно додані!")
    except Exception as e:
        print(f" -> [ПОМИЛКА Geocoding]: {e}")
        
    return shops

# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------
def main():
    print("=" * 60)
    print("ЗБІР ДАНИХ МАГАЗИНІВ МОБІЛЬНИХ ОПЕРАТОРІВ (З ЗАХИСТОМ ТА БЕКАПОМ)")
    print("=" * 60)

    # Step 0: Робимо бекап існуючих результатів перед оновленням
    make_initial_backup()

    # Step 1-3: Отримуємо дані (з підтримкою автоматичного фоллбеку на бекап при падінні)
    lifecell_shops = fetch_lifecell()
    vodafone_shops = fetch_vodafone()
    kyivstar_shops = fetch_kyivstar()

    all_shops = lifecell_shops + vodafone_shops + kyivstar_shops

    if not all_shops:
        print("\n[КРИТИЧНО] Жоден провайдер не повернув даних і бекапи відсутні. Скасування збереження.")
        return

    # Збагачуємо всі записи областю за координатами
    all_shops = enrich_with_regions(all_shops)

    # Збереження у JSON
    try:
        with open(OUTPUT_JSON_FILE, "w", encoding="utf-8") as f:
            json.dump(all_shops, f, ensure_ascii=False, indent=2)
        print(f"\n[УСПІХ] Збережено JSON: {OUTPUT_JSON_FILE}")
    except Exception as e:
        print(f"\n[ПОМИЛКА] Не вдалося зберегти JSON: {e}")

    # Збереження в Excel
    try:
        df = pd.DataFrame(all_shops)
        df.rename(columns={
            "provider": "Оператор",
            "id": "ID",
            "name": "Назва/Орієнтир",
            "address": "Адреса",
            "city": "Місто",
            "region": "Область (оригінал)",
            "geo_region": "Область (визначена)",
            "lat": "Широта",
            "lng": "Довгота",
            "working_hours": "Графік роботи"
        }, inplace=True)
        
        df.to_excel(OUTPUT_EXCEL_FILE, index=False)
        print(f"[УСПІХ] Збережено Excel: {OUTPUT_EXCEL_FILE}")
    except Exception as e:
        print(f"[ПОМИЛКА] Не вдалося зберегти Excel: {e}")

    print("\nПідсумок:")
    print(f" - lifecell: {len(lifecell_shops)}")
    print(f" - Vodafone: {len(vodafone_shops)}")
    print(f" - Kyivstar: {len(kyivstar_shops)}")
    print(f" - Всього:   {len(all_shops)}")

if __name__ == "__main__":
    main()
