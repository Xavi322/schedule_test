import os
import json
import zlib
import struct
from downloader import sync_cloud_files
from parser import generate_week_calendar, BASE_DIR

CLOUD_FOLDER_URL = "https://cloud.mail.ru/public/s9pr/NyGVkqDRw"

def create_calendar_png(filepath, size=512):
    """Генерирует 100% валидную PNG-иконку календаря."""
    rows = []
    for y in range(size):
        # ВАЖНО: bytearray(1) создает стартовый нулевой байт фильтра PNG!
        row = bytearray(1)
        ny = y / size
        for x in range(size):
            nx = x / size
            # Синий фон (#3b82f6)
            r, g, b = 59, 130, 246

            # Белая карточка календаря по центру
            if 0.18 <= nx <= 0.82 and 0.22 <= ny <= 0.82:
                if ny <= 0.36:
                    # Темная шапка (#1e293b)
                    r, g, b = 30, 41, 59
                else:
                    # Белое поле
                    r, g, b = 255, 255, 255
                    # Сетка пар
                    for cx in (0.33, 0.50, 0.67):
                        for cy in (0.50, 0.67):
                            if abs(nx - cx) < 0.055 and abs(ny - cy) < 0.055:
                                if cx > 0.6 and cy > 0.6:
                                    r, g, b = 234, 88, 12
                                else:
                                    r, g, b = 59, 130, 246
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
    """Создает чистые PNG-иконки для лаунчера Android и Firefox."""
    create_calendar_png(os.path.join(BASE_DIR, "app-icon-192.png"), size=192)
    create_calendar_png(os.path.join(BASE_DIR, "app-icon-512.png"), size=512)
    create_calendar_png(os.path.join(BASE_DIR, "icon.png"), size=192)

    manifest = {
        "name": "Расписание занятий",
        "short_name": "Расписание",
        "start_url": "./index.html",
        "display": "standalone",
        "background_color": "#3b82f6",
        "theme_color": "#1e293b",
        "icons": [
            {
                "src": "app-icon-192.png",
                "sizes": "192x192",
                "type": "image/png"
            },
            {
                "src": "app-icon-512.png",
                "sizes": "512x512",
                "type": "image/png"
            }
        ]
    }
    with open(os.path.join(BASE_DIR, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

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
    
    print("=== 3. Создаем файлы PWA ===")
    generate_pwa_files()
    print("🚀 Все файлы готовы!")
