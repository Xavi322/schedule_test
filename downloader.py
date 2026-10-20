import os
import re
import json
import datetime
import requests
from urllib.parse import quote

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Referer": "https://mail.ru"
}


def extract_folder_data_from_html(session, public_url):
    """Загружает страницу папки Mail.ru и вытаскивает точный шлюз скачивания и список таблиц."""
    resp = session.get(public_url, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    html = resp.text

    # 1. Ищем точный URL шлюза скачивания в секции "weblink_get" (независимо от порядка полей)
    base_dl_url = ""
    wg_pos = html.find('"weblink_get"')
    if wg_pos != -1:
        snippet = html[wg_pos:wg_pos + 600]
        url_m = re.search(r'"url"\s*:\s*("https?:[^"]+")', snippet)
        if url_m:
            base_dl_url = json.loads(url_m.group(1))

    # Если вдруг не нашли по тексту, ищем любой шлюз cloclo / datacloudmail для weblink/get
    if not base_dl_url:
        fallback_m = re.search(r'(https?://[a-zA-Z0-9.-]+/(?:weblink/get|public/get))', html.replace("\\/", "/"))
        if fallback_m:
            base_dl_url = fallback_m.group(1)

    # 2. Собираем все таблицы (.xlsx, .xls) из папки
    items = []
    seen_weblinks = set()

    for m in re.finditer(r'"weblink"\s*:\s*("[^"]+")', html):
        try:
            weblink_val = json.loads(m.group(1)).strip("/")
        except Exception:
            continue

        if weblink_val in seen_weblinks:
            continue

        _, _, filename = weblink_val.rpartition("/")
        if not filename:
            continue

        low_name = filename.lower()
        if low_name.endswith((".xlsx", ".xls")) or "ссылк" in low_name:
            seen_weblinks.add(weblink_val)
            items.append({"name": filename, "weblink": weblink_val})

    return items, base_dl_url


def is_current_week_file(filename, today=None):
    """
    Понимает оба формата дат в названии файла:
    1) Сокращенный: '19-25.10.2026' (где месяц указан один раз в конце)
    2) Полный: '19.10 - 25.10.2026' или '28.10-03.11.2026'
    """
    if today is None:
        today = datetime.date.today()

    monday = today - datetime.timedelta(days=today.weekday())
    sunday = monday + datetime.timedelta(days=6)

    # Вариант А: Формат '19-25.10.2026' или '19-25.10'
    short_m = re.search(r"(\d{1,2})\s*[-–—]\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?", filename)
    if short_m:
        d1, d2, month, y_str = int(short_m.group(1)), int(short_m.group(2)), int(short_m.group(3)), short_m.group(4)
        year = int(y_str) if y_str else today.year
        if year < 100:
            year += 2000
        try:
            start_d = datetime.date(year, month, d1)
            end_d = datetime.date(year, month, d2)
            return start_d <= today <= end_d or start_d == monday
        except ValueError:
            pass

    # Вариант Б: Две отдельные даты с точками ('19.10 - 25.10' или переход месяца '26.10 - 01.11')
    matches = re.findall(r"(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?", filename)
    parsed_dates = []
    for d_str, m_str, y_str in matches:
        day, month = int(d_str), int(m_str)
        year = int(y_str) if y_str else today.year
        if year < 100:
            year += 2000
        try:
            parsed_dates.append(datetime.date(year, month, day))
        except ValueError:
            continue

    if len(parsed_dates) >= 2:
        start_d, end_d = parsed_dates[:2]
        if end_d < start_d:
            end_d = end_d.replace(year=start_d.year + 1)
        return start_d <= today <= end_d or start_d == monday
    elif len(parsed_dates) == 1:
        first_date = next(iter(parsed_dates))
        return monday <= first_date <= sunday

    return False


def save_item(session, base_dl_url, item, local_filename):
    """Скачивает файл через найденный шлюз и проверяет ZIP/Excel сигнатуру."""
    weblink = item.get("weblink", "").strip("/")
    clean_base = base_dl_url.rstrip("/")
    file_url = clean_base + "/" + quote(weblink, safe="/")

    print(f"⏳ Качаем '{item.get('name')}' -> {local_filename}")
    print(f"   Запрос к: {file_url}")
    resp = session.get(file_url, headers=HEADERS, timeout=30)
    resp.raise_for_status()

    content = resp.content
    if not (content.startswith(b"PK") or content.startswith(b"\xd0\xcf")):
        print(f"❌ ОШИБКА: Сервер вернул не Excel-файл (код {resp.status_code}, размер: {len(content)} байт)!")
        print("   URL после редиректа:", resp.url)
        return False

    save_path = os.path.join(BASE_DIR, local_filename)
    with open(save_path, "wb") as f:
        f.write(content)
    print(f"✅ Успешно сохранено ({len(content) // 1024} КБ): {save_path}\n")
    return True


def sync_cloud_files(folder_url):
    session = requests.Session()
    print("🌐 Открываем страницу папки Облака Mail.ru...")
    xlsx_items, base_dl_url = extract_folder_data_from_html(session, folder_url)
    print(f"🔗 Найденный шлюз скачивания: {base_dl_url}\n")

    if not xlsx_items:
        print("❌ На странице не найдено .xlsx файлов!")
        return

    print("📂 Найдены таблицы в облаке:")
    for i in xlsx_items:
        print(f"   • {i.get('name')}")
    print()

    # 1. Качаем файл со ссылками
    links_item = next((i for i in xlsx_items if "ссылк" in i.get("name", "").lower()), None)
    if links_item:
        save_item(session, base_dl_url, links_item, "links.xlsx")
    else:
        print("⚠️ Файл со словом 'Ссылки' не найден.")

    # 2. Качаем расписание на ТЕКУЩУЮ неделю
    sched_candidates = [i for i in xlsx_items if i != links_item]
    current_sched = next((i for i in sched_candidates if is_current_week_file(i.get("name", ""))), None)

    if current_sched:
        save_item(session, base_dl_url, current_sched, "schedule.xlsx")
    else:
        print("⚠️ Не удалось определить текущую неделю по датам в названии!")


if __name__ == "__main__":
    CLOUD_FOLDER_URL = "https://cloud.mail.ru/public/s9pr/NyGVkqDRw"
    sync_cloud_files(CLOUD_FOLDER_URL)
