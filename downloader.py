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

def get_moscow_today():
    """Возвращает сегодняшнюю дату по Москве (UTC+3), чтобы сервер GitHub не путал дни."""
    tz_msk = datetime.timezone(datetime.timedelta(hours=3))
    return datetime.datetime.now(tz_msk).date()

def extract_folder_data_from_html(session, public_url):
    resp = session.get(public_url, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    html = resp.text

    base_dl_url = ""
    wg_pos = html.find('"weblink_get"')
    if wg_pos != -1:
        snippet = html[wg_pos:wg_pos + 600]
        url_m = re.search(r'"url"\s*:\s*("https?:[^"]+")', snippet)
        if url_m:
            base_dl_url = json.loads(url_m.group(1))

    if not base_dl_url:
        fallback_m = re.search(r'(https?://[a-zA-Z0-9.-]+/(?:weblink/get|public/get))', html.replace("\\/", "/"))
        if fallback_m:
            base_dl_url = fallback_m.group(1)

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

def parse_week_range(filename, today):
    """
    Извлекает даты начала и конца недели из названия файла.
    Возвращает кортеж (start_date, end_date, label) или None.
    """
    # Вариант А: '19-25.10.2026' или '19-25.10'
    short_m = re.search(r"(\d{1,2})\s*[-–—]\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?", filename)
    if short_m:
        d1, d2, month, y_str = int(short_m.group(1)), int(short_m.group(2)), int(short_m.group(3)), short_m.group(4)
        year = int(y_str) if y_str else today.year
        if year < 100:
            year += 2000
        try:
            s_d = datetime.date(year, month, d1)
            e_d = datetime.date(year, month, d2)
            label = f"{s_d.strftime('%d.%m')} — {e_d.strftime('%d.%m')}"
            return s_d, e_d, label
        except ValueError:
            pass

    # Вариант Б: '26.10 - 01.11.2026'
    matches = re.findall(r"(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?", filename)
    parsed = []
    for d_str, m_str, y_str in matches:
        day, month = int(d_str), int(m_str)
        year = int(y_str) if y_str else today.year
        if year < 100:
            year += 2000
        try:
            parsed.append(datetime.date(year, month, day))
        except ValueError:
            continue

    if len(parsed) >= 2:
        s_d, e_d = parsed[:2]
        if e_d < s_d:
            e_d = e_d.replace(year=s_d.year + 1)
        label = f"{s_d.strftime('%d.%m')} — {e_d.strftime('%d.%m')}"
        return s_d, e_d, label
    elif len(parsed) == 1:
        s_d = next(iter(parsed))
        e_d = s_d + datetime.timedelta(days=6)
        label = f"{s_d.strftime('%d.%m')} — {e_d.strftime('%d.%m')}"
        return s_d, e_d, label

    return None

def save_item(session, base_dl_url, item, local_filename):
    weblink = item.get("weblink", "").strip("/")
    file_url = base_dl_url.rstrip("/") + "/" + quote(weblink, safe="/")
    print(f"⏳ Качаем '{item.get('name')}' -> {local_filename}")
    resp = session.get(file_url, headers=HEADERS, timeout=30)
    resp.raise_for_status()

    content = resp.content
    if not (content.startswith(b"PK") or content.startswith(b"\xd0\xcf")):
        print(f"❌ ОШИБКА: Сервер вернул не Excel-файл!")
        return False

    save_path = os.path.join(BASE_DIR, local_filename)
    with open(save_path, "wb") as f:
        f.write(content)
    print(f"✅ Сохранено: {local_filename}")
    return True

def sync_cloud_files(folder_url):
    session = requests.Session()
    today = get_moscow_today()
    monday = today - datetime.timedelta(days=today.weekday())

    print("🌐 Проверяем папку Облака Mail.ru...")
    xlsx_items, base_dl_url = extract_folder_data_from_html(session, folder_url)

    if not xlsx_items:
        print("❌ На странице не найдено .xlsx файлов!")
        return []

    # 1. Качаем файл со ссылками
    links_item = next((i for i in xlsx_items if "ссылк" in i.get("name", "").lower()), None)
    if links_item:
        save_item(session, base_dl_url, links_item, "links.xlsx")

    # 2. Собираем ВСЕ файлы расписания (текущую и будущие недели)
    sched_candidates = [i for i in xlsx_items if i != links_item]
    dated_weeks = []

    for item in sched_candidates:
        info = parse_week_range(item.get("name", ""), today)
        if info:
            s_d, e_d, label = info
            # Берем недели, которые еще не закончились (или начались на этой неделе)
            if e_d >= monday:
                is_curr = (s_d <= today <= e_d) or (s_d == monday)
                dated_weeks.append({
                    "start": s_d,
                    "item": item,
                    "label": f"📅 {label} (Текущая)" if is_curr else f"📅 {label}",
                    "is_current": is_curr
                })

    # Если вдруг в папке лежат только старые недели — берем их все, чтобы сайт не был пустым
    if not dated_weeks:
        for item in sched_candidates:
            info = parse_week_range(item.get("name", ""), today)
            lbl = f"📅 {info[2]}" if info else item.get("name", "Неделя")
            s_d = info[0] if info else today
            dated_weeks.append({"start": s_d, "item": item, "label": lbl, "is_current": False})

    # Сортируем недели строго по хронологии (от ранней к более поздней)
    dated_weeks.sort(key=lambda w: w["start"])

    weeks_meta = []
    for idx, w in enumerate(dated_weeks):
        local_fname = f"schedule_{idx}.xlsx"
        if save_item(session, base_dl_url, w["item"], local_fname):
            weeks_meta.append({
                "id": f"week_{idx}",
                "label": w["label"],
                "file": local_fname,
                "is_current": w["is_current"]
            })
            # Для обратной совместимости копируем первую/текущую неделю также в schedule.xlsx
            if w["is_current"] or idx == 0:
                save_item(session, base_dl_url, w["item"], "schedule.xlsx")

    meta_path = os.path.join(BASE_DIR, "weeks_meta.json")
    with open(meta_path, "w", encoding="utf-8") as mf:
        json.dump(weeks_meta, mf, ensure_ascii=False, indent=2)

    print(f"🎉 Загружено недель расписания: {len(weeks_meta)}")
    return weeks_meta

if __name__ == "__main__":
    CLOUD_FOLDER_URL = "https://mail.rupublic/s9pr/NyGVkqDRw"
    sync_cloud_files(CLOUD_FOLDER_URL)
