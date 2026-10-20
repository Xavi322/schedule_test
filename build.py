import os
import json
from downloader import sync_cloud_files
from parser import generate_week_calendar, BASE_DIR

CLOUD_FOLDER_URL = "https://cloud.mail.ru/public/s9pr/NyGVkqDRw"

def generate_pwa_files():
    """Создает локальную иконку, manifest.json и Service Worker для режима без адресной строки."""
    # 1. Локальная векторная иконка календаря (чтобы браузер не блокировал чужие сайты)
    svg_icon = (
        '<svg xmlns="http://w3.org" viewBox="0 0 512 512" width="512" height="512">'
        '<rect width="512" height="512" rx="115" fill="#3b82f6"/>'
        '<rect x="76" y="116" width="360" height="320" rx="40" fill="#ffffff"/>'
        '<path d="M76 156 a40 40 0 0 1 40 -40 h280 a40 40 0 0 1 40 40 v50 h-360 z" fill="#1e293b"/>'
        '<circle cx="176" cy="280" r="24" fill="#3b82f6"/>'
        '<circle cx="256" cy="280" r="24" fill="#3b82f6"/>'
        '<circle cx="336" cy="280" r="24" fill="#3b82f6"/>'
        '<circle cx="176" cy="360" r="24" fill="#3b82f6"/>'
        '<circle cx="256" cy="360" r="24" fill="#ea580c"/>'
        '</svg>'
    )
    with open(os.path.join(BASE_DIR, "icon.svg"), "w", encoding="utf-8") as f:
        f.write(svg_icon)

    # 2. Правильный манифест со ссылками на ЛОКАЛЬНУЮ иконку
    manifest = {
        "id": "./",
        "name": "Расписание занятий",
        "short_name": "Расписание",
        "start_url": "./",
        "scope": "./",
        "display": "standalone",
        "background_color": "#f4f6f9",
        "theme_color": "#1e293b",
        "icons": [
            {
                "src": "icon.svg",
                "sizes": "192x192 512x512",
                "type": "image/svg+xml",
                "purpose": "any maskable"
            }
        ]
    }
    with open(os.path.join(BASE_DIR, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    # 3. Обязательный Service Worker (без него Firefox и Chrome не прячут адресную строку)
    sw_code = (
        "self.addEventListener('install', function(e) { self.skipWaiting(); });\n"
        "self.addEventListener('activate', function(e) { e.waitUntil(self.clients.claim()); });\n"
        "self.addEventListener('fetch', function(e) { e.respondWith(fetch(e.request)); });\n"
    )
    with open(os.path.join(BASE_DIR, "sw.js"), "w", encoding="utf-8") as f:
        f.write(sw_code)

if __name__ == "__main__":
    print("=== 1. Скачиваем таблицы из Облака Mail.ru ===")
    sync_cloud_files(CLOUD_FOLDER_URL)
    
    print("=== 2. Генерируем веб-страницу index.html ===")
    generate_week_calendar("schedule.xlsx", "links.xlsx", "index.html")
    
    print("=== 3. Создаем файлы PWA (icon.svg, manifest.json, sw.js) ===")
    generate_pwa_files()
    print("🚀 Все файлы готовы!")
