import json
import os
from curl_cffi import requests

BACKUP_JSON_FILE = "all_shops_bcp.json"

def fetch_and_save_lifecell():
    print("[LIFECELL] Отримуємо магазини з API lifecell...")
    url = "https://www.lifecell.ua/location-services/api/v1/pos/?limit=50000&offset=0&type=LIFECELL"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "uk-UA,uk;q=0.9,en-US;q=0.8,en;q=0.7",
        "Sec-Ch-Ua": '"Google Chrome";v="123", "Not:A-Brand";v="8", "Chromium";v="123"',
        "Sec-Ch-Ua-Mobile": "?0",
        "Sec-Ch-Ua-Platform": '"Windows"',
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
        "Referer": "https://www.lifecell.ua/uk/shops/"
    }

    try:
        response = requests.get(url, headers=headers, impersonate="chrome123", timeout=25)
        
        if response.status_code == 200 and "application/json" in response.headers.get("Content-Type", ""):
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

            if normalized:
                with open(BACKUP_JSON_FILE, "w", encoding="utf-8") as f:
                    json.dump(normalized, f, ensure_ascii=False, indent=2)
                print(f"[УСПІХ] Збережено {len(normalized)} магазинів lifecell у файл {BACKUP_JSON_FILE}")
                return

        print(f"[УВАГА] Сервер lifecell повернув статус {response.status_code} або не-JSON відповідь.")

    except Exception as e:
        print(f"[ПОМИЛКА] Запит до lifecell впав: {e}")

if __name__ == "__main__":
    fetch_and_save_lifecell()
