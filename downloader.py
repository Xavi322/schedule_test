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
    """Делает до 3 попыток получить список файлов (защита от пустых ответов Mail.ru)."""
    for attempt in range(1, 4):
        try:
            # Добавляем случайный параметр, чтобы обходить кэш Mail.ru
            sep = "&" if "?" in public_url else "?"
            url_nocache = f"{public_url}{sep}_={int(time.time())}"
            resp = session.get(url_nocache, headers=HEADERS, timeout=15)
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
                fb = re.search(r'(https?://[a-zA-Z0-9.-]+/(?:weblink/get|public/get))', html.replace("\\/", "/"))
                if fb:
                    base_dl_url = fb.group(1)

            items = []
            seen = set()
            for m in re.finditer(r'"weblink"\s*:\s*("[^"]+")', html):
                try:
                    weblink_val = json.loads(m.group(1)).strip("/")
                except Exception:
                    continue
                if weblink_val in seen:
                    continue
                _, _, filename = weblink_val.rpartition("/")
                if not filename:
                    continue
                low_name = filename.lower()
                # Берем таблицы (.xlsx, .xls) И файлы .pdf (если в них есть даты или слово 'распис')
                if low_name.endswith((".xlsx", ".xls", ".pdf")) or "ссылк" in low_name:
                    seen.add(weblink_val)
                    items.append({"name": filename, "weblink": weblink_val})

            if items:
                return items, base_dl_url
            print(f"⚠️ Попытка {attempt}/3: Mail.ru вернул пустой список, пробуем снова через 2 сек...")
        except Exception as e:
            print(f"⚠️ Попытка {attempt}/3 не удалась ({e}), ждем 2 сек...")
        time.sleep(2)

    return [], ""


def parse_week_range(filename, today):
    """
    Всеядный парсер дат из названия файла.
    Понимает: '09-15.11.2026', '09.11 - 15.11', 'с 09.11 по 15.11', '9-15 ноября' и т.д.
    """
    low = filename.lower()
    months_map = {
        "янв": 1, "фев": 2, "мар": 3, "апр": 4, "мая": 5, "май": 5, "июн": 6,
        "июл": 7, "авг": 8, "сен": 9, "окт": 10, "ноя": 11, "дек": 12
    }

    # 1. Ищем две полные числовые даты: '09.11 ... 15.11' (через тире, 'по' или пробелы)
    full_m = re.search(
        r"(\d{1,2})\s*\.\s*(\d{1,2})(?:\s*\.\s*(\d{2,4}))?\s*(?:[-–—]+|по)\s*(\d{1,2})\s*\.\s*(\d{1,2})(?:\s*\.\s*(\d{2,4}))?",
        low
    )
    if full_m:
        d1, m1, y1 = int(full_m.group(1)), int(full_m.group(2)), full_m.group(3)
        d2, m2, y2 = int(full_m.group(4)), int(full_m.group(5)), full_m.group(6)
        year2 = int(y2) if y2 else (int(y1) if y1 else today.year)
        if year2 < 100: year2 += 2000
        year1 = int(y1) if y1 else year2
        if year1 < 100: year1 += 2000
        if not y1 and m1 == 12 and m2 == 1: year1 = year2 - 1
        try:
            s_d = datetime.date(year1, m1, d1)
            e_d = datetime.date(year2, m2, d2)
            return s_d, e_d, f"{s_d.strftime('%d.%m')} — {e_d.strftime('%d.%m')}"
        except ValueError:
            pass

    # 2. Сокращенный диапазон внутри одного месяца: '09-15.11.2026' или 'с 9 по 15.11'
    short_m = re.search(
        r"(?<![.\d])(\d{1,2})\s*(?:[-–—]+|по)\s*(\d{1,2})\s*\.\s*(\d{1,2})(?:\s*\.\s*(\d{2,4}))?",
        low
    )
    if short_m:
        d1, d2, month, y_str = int(short_m.group(1)), int(short_m.group(2)), int(short_m.group(3)), short_m.group(4)
        year = int(y_str) if y_str else today.year
        if year < 100: year += 2000
        try:
            s_d = datetime.date(year, month, d1)
            e_d = datetime.date(year, month, d2)
            return s_d, e_d, f"{s_d.strftime('%d.%m')} — {e_d.strftime('%d.%m')}"
        except ValueError:
            pass

    # 3. Если месяц написан словом ('9-15 ноября' или 'с 9 по 15 ноября')
    for m_prefix, m_num in months_map.items():
        if m_prefix in low:
            word_m = re.search(r"(\d{1,2})\s*(?:[-–—]+|по)\s*(\d{1,2})\s*" + m_prefix, low)
            if word_m:
                d1, d2 = int(word_m.group(1)), int(word_m.group(2))
                try:
                    s_d = datetime.date(today.year, m_num, d1)
                    e_d = datetime.date(today.year, m_num, d2)
                    return s_d, e_d, f"{s_d.strftime('%d.%m')} — {e_d.strftime('%d.%m')}"
                except ValueError:
                    pass

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

def sync_single_folder(session, folder_url, prefix, today, monday):
    """Скачивает ссылки и все актуальные недели из папки без потери нестандартных файлов."""
    print(f"\n🌐 Проверяем папку [{prefix.upper()}]: {folder_url}")
    xlsx_items, base_dl_url = extract_folder_data_from_html(session, folder_url)
    if not xlsx_items:
        print(f"⚠️ В папке [{prefix}] не найдено таблиц!")
        return []

    print(f"📂 Все найденные таблицы в [{prefix}]: {[i.get('name') for i in xlsx_items]}")

    # 1. Качаем файл со ссылками
    links_item = next((i for i in xlsx_items if "ссылк" in i.get("name", "").lower()), None)
    if links_item:
        save_item(session, base_dl_url, links_item, f"{prefix}_links.xlsx")
        if prefix == "main":
            save_item(session, base_dl_url, links_item, "links.xlsx")

    # 2. Разбираем все остальные таблицы расписания
    sched_candidates = [i for i in xlsx_items if i != links_item]
    dated_weeks = []
    undated_weeks = []

    for item in sched_candidates:
        raw_name = item.get("name", "")
        info = parse_week_range(raw_name, today)
        if info:
            s_d, e_d, label = info
            if e_d >= monday:
                is_curr = (s_d <= today <= e_d + datetime.timedelta(days=1)) or (s_d == monday)
                dated_weeks.append({
                    "start": s_d, "item": item, "is_current": is_curr,
                    "label": f"📅 {label} (Текущая)" if is_curr else f"📅 {label}"
                })
        else:
            clean_lbl = re.sub(r"\.xlsx?$", "", raw_name, flags=re.IGNORECASE).strip()
            undated_weeks.append({
                "start": today, "item": item, "is_current": False,
                "label": f"📅 {clean_lbl}"
            })
    if not dated_weeks and not undated_weeks:
        for item in sched_candidates:
            info = parse_week_range(item.get("name", ""), today)
            s_d, _, label = info if info else (today, today, item.get("name", "Неделя"))
            dated_weeks.append({"start": s_d, "item": item, "label": f"📅 {label}", "is_current": False})

    dated_weeks.sort(key=lambda w: w["start"])
    final_weeks = dated_weeks + undated_weeks

    # Если ни одна неделя не помечена как текущая, делаем текущей самую первую
    if final_weeks and not any(w["is_current"] for w in final_weeks):
        first_w = next(iter(final_weeks))
        first_w["is_current"] = True

    folder_meta = []
    for idx, w in enumerate(final_weeks):
        local_fname = f"{prefix}_schedule_{idx}.xlsx"
        if save_item(session, base_dl_url, w["item"], local_fname):
            folder_meta.append({
                "id": f"{prefix}_week_{idx}",
                "label": w["label"],
                "file": local_fname,
                "is_current": w["is_current"]
            })
    return folder_meta


def sync_cloud_files(folders_config):
    """
    Принимает либо одну строку-ссылку, либо словарь вида:
    {"main": {"title": "Очное", "url": "..."}, "iup": {"title": "ИУП / Заочное", "url": "..."}}
    """
    session = requests.Session()
    today = get_moscow_today()
    monday = today - datetime.timedelta(days=today.weekday())

    if isinstance(folders_config, str):
        folders_config = {"main": {"title": "Основное", "url": folders_config}}

    all_meta = {}
    for key, cfg in folders_config.items():
        url = cfg.get("url", "").strip()
        if not url or "XXXX" in url:
            continue
        weeks_list = sync_single_folder(session, url, key, today, monday)
        if weeks_list:
            all_meta[key] = {
                "title": cfg.get("title", key),
                "links_file": f"{key}_links.xlsx",
                "weeks": weeks_list
            }

    meta_path = os.path.join(BASE_DIR, "weeks_meta.json")
    with open(meta_path, "w", encoding="utf-8") as mf:
        json.dump(all_meta, mf, ensure_ascii=False, indent=2)

    print(f"\n🎉 Синхронизация завершена! Потоков загружено: {len(all_meta)}")
    return all_meta

if __name__ == "__main__":
    FOLDERS = {
        "main": {
            "title": "🎓 Очное",
            "url": "https://cloud.mail.ru/public/s9pr/NyGVkqDRw"
        },
        "iup": {
            "title": "💼 ИУП (Заочное)",
            "url": "https://cloud.mail.ru/public/ycXv/dTTxtDU6Z"
        }
    }
    sync_cloud_files(FOLDERS)
