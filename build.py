import os
import json
import zlib
import struct
from downloader import sync_cloud_files
from parser import generate_week_calendar, BASE_DIR

CLOUD_FOLDER_URL = "https://cloud.mail.ru/public/s9pr/NyGVkqDRw"

def create_calendar_png(filepath, size=512):
    """Генерирует настоящую PNG-иконку календаря чистым Python (без сторонних библиотек)."""
    rows = []
    for y in range(size):
        row = bytearray([0])  # Байт фильтра PNG в начале каждой строки
        ny = y / size
        for x in range(size):
            nx = x / size
            # 1. Базовый фон — приятный синий (#3b82f6)
            r, g, b = 59, 130, 246

            # 2. Белая карточка календаря по центру
            if 0.18 <= nx <= 0.82 and 0.22 <= ny <= 0.82:
                if ny <= 0.36:
                    # Темная шапка календаря (#1e293b)
                    r, g, b = 30, 41, 59
                else:
                    # Белое поле календаря
                    r, g, b = 255, 255, 255
                    # Рисуем 6 аккуратных цветных квадратиков (пары в сетке)
                    for cx in (0.33, 0.50, 0.67):
                        for cy in (0.50, 0.67):
                            if abs(nx - cx) < 0.055 and abs(ny - cy) < 0.055:
                                if cx > 0.6 and cy > 0.6:
                                    r, g, b = 234, 88, 12  # Оранжевый акцент на одной паре
                                else:
                                    r, g, b = 59, 130, 246 # Синие слоты
            row.extend((r, g, b, 255))
        rows.append(bytes(row))

    raw_data = b"".join(rows)

    def make_chunk(chunk_type, data):
        crc = zlib.crc32(chunk_type + data) & 0xffffffff
        return struct.pack(">I", len(data)) + chunk_type + data + struct.pack(">I", crc)

    ihdr = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    png_bytes = (
        b"\x89PNG\r\n\x1a\n"
        + make_chunk(b"IHDR", ihdr)
        + make_chunk(b"IDAT", zlib.compress(raw_data, 9))
        + make_chunk(b"IEND", b"")
    )
    with open(filepath, "wb") as f:
        f.write(png_bytes)

def generate_pwa_files():
    """Создает icon.png, manifest.json и sw.js для Android."""
    # 1. Генерируем растровую PNG-иконку (Android понимает только PNG на рабочем столе!)
    create_calendar_png(os.path.join(BASE_DIR, "icon.png"), size=512)

    # 2. В манифесте указываем icon.png и разделяем 'any' и 'maskable'
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
                "src": "icon.png",
                "sizes": "192x192 512x512",
                "type": "image/png",
                "purpose": "any"
            },
            {
                "src": "icon.png",
                "sizes": "192x192 512x512",
                "type": "image/png",
                "purpose": "maskable"
            }
        ]
    }
    with open(os.path.join(BASE_DIR, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    # 3. Service Worker
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
    
    print("=== 3. Создаем файлы PWA (icon.png, manifest.json, sw.js) ===")
    generate_pwa_files()
    print("🚀 Все файлы готовы!")
