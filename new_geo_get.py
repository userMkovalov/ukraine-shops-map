import json
import requests
import pandas as pd
from bs4 import BeautifulSoup
import reverse_geocoder as rg

# Налаштування
OUTPUT_JSON_FILE = "all_shops.json"
OUTPUT_EXCEL_FILE = "all_shops.xlsx"

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
# 1. LIFECELL
# ------------------------------------------------------------
def fetch_lifecell():
    print("[1/3] Завантажуємо магазини lifecell...")
    url = "https://www.lifecell.ua/location-services/api/v1/pos/?limit=50000&offset=0&type=LIFECELL"
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()

        raw = data.get("results", [])
        normalized = []

        for item in raw:
            lat = item.get("lat") or item.get("latitude")
            lng = item.get("lng") or item.get("lon") or item.get("longitude")

            if lat is None or lng is None:
                continue

            try:
                lat = float(lat)
                lng = float(lng)
            except (TypeError, ValueError):
                continue

            normalized.append({
                "provider": "lifecell",
                "id": item.get("id"),
                "name": item.get("name") or "lifecell",
                "address": item.get("address") or "",
                "city": item.get("city") or "",
                "region": item.get("region") or "",
                "lat": lat,
                "lng": lng,
                "working_hours": item.get("working_hours") or item.get("schedule") or ""
            })

        print(f" -> Знайдено {len(normalized)} магазинів lifecell.")
        return normalized

    except Exception as e:
        print(f" -> Помилка при отриманні lifecell: {e}")
        return []

# ------------------------------------------------------------
# 2. VODAFONE
# ------------------------------------------------------------
def fetch_vodafone():
    print("[2/3] Завантажуємо магазини Vodafone...")
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
            print(" -> Помилка: ng-state не знайдено або він порожній.")
            return []

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

        # Видалення дублікатів
        unique = []
        seen = set()
        for shop in normalized:
            key = (shop["provider"], shop["lat"], shop["lng"], shop["address"])
            if key not in seen:
                seen.add(key)
                unique.append(shop)

        print(f" -> Знайдено {len(unique)} магазинів Vodafone.")
        return unique

    except Exception as e:
        print(f" -> Помилка при отриманні Vodafone: {e}")
        return []

# ------------------------------------------------------------
# 3. KYIVSTAR
# ------------------------------------------------------------
def fetch_kyivstar():
    print("[3/3] Завантажуємо магазини Kyivstar...")
    url = "https://kyivstar.ua/api/pos-shops?locale=uk&start=0&limit=10000"
    
    # Повний набір заголовків, щоб запобігти таймаутам/блокуванню
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

            # Збирання адреси
            street_type = item.get("streetType", "") or ""
            street_name = item.get("streetName", "") or ""
            building = item.get("buildingNumber", "") or ""
            full_address = f"{street_type} {street_name} {building}".strip()

            # Місто
            city_obj = item.get("city", {}) or {}
            city_name = city_obj.get("city", "") if isinstance(city_obj, dict) else ""

            # Графік роботи
            wh_list = item.get("workingHours", [])
            wh_str_list = []
            if isinstance(wh_list, list):
                for wh in wh_list:
                    if isinstance(wh, dict):
                        day = wh.get("dayOfWeek", "")
                        start = wh.get("startTime", "")[:5]
                        end = wh.get("endTime", "")[:5]
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
                "region": "",  # Kyivstar API не повертає область, додасть reverse_geocoder
                "lat": lat,
                "lng": lng,
                "working_hours": working_hours
            })

        print(f" -> Знайдено {len(normalized)} магазинів Kyivstar.")
        return normalized

    except Exception as e:
        print(f" -> Помилка при отриманні Kyivstar: {e}")
        return []

# ------------------------------------------------------------
# 4. REVERSE GEOCODING (Визначення області за координатами)
# ------------------------------------------------------------
def enrich_with_regions(shops):
    print("\n[Геолокація] Визначаємо області за координатами через reverse_geocoder...")
    
    coords = [(shop["lat"], shop["lng"]) for shop in shops]
    
    if not coords:
        return shops

    results = rg.search(coords)

    for i, shop in enumerate(shops):
        res = results[i]
        admin1 = res.get("admin1", "").replace(" Oblast", "").strip()
        ua_region = REGION_MAP.get(admin1, admin1)
        
        shop["geo_region"] = ua_region
        if not shop["region"]:
            shop["region"] = ua_region

    print(" -> Області успішно додані!")
    return shops

# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------
def main():
    print("=" * 60)
    print("ЗБІР ДАНИХ МАГАЗИНІВ Мобільних Операторів")
    print("=" * 60)

    lifecell_shops = fetch_lifecell()
    vodafone_shops = fetch_vodafone()
    kyivstar_shops = fetch_kyivstar()

    all_shops = lifecell_shops + vodafone_shops + kyivstar_shops

    # Збагачуємо дані областю за координатами
    all_shops = enrich_with_regions(all_shops)

    # Збереження у JSON
    with open(OUTPUT_JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(all_shops, f, ensure_ascii=False, indent=2)
    print(f"\n[УСПІХ] Збережено JSON: {OUTPUT_JSON_FILE}")

    # Збереження в Excel
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

    print("\nПідсумок:")
    print(f" - lifecell: {len(lifecell_shops)}")
    print(f" - Vodafone: {len(vodafone_shops)}")
    print(f" - Kyivstar: {len(kyivstar_shops)}")
    print(f" - Всього:   {len(all_shops)}")

if __name__ == "__main__":
    main()
