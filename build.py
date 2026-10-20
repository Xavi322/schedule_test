from downloader import sync_cloud_files
from parser import generate_week_calendar

CLOUD_FOLDER_URL = "https://cloud.mail.ru/public/s9pr/NyGVkqDRw"

if __name__ == "__main__":
    print("=== 1. Скачиваем таблицы из Облака Mail.ru ===")
    sync_cloud_files(CLOUD_FOLDER_URL)

    print("=== 2. Генерируем веб-страницу index.html ===")
    generate_week_calendar(
        filename="schedule.xlsx",
        links_file="links.xlsx",
        output_html="index.html"
    )
    print("🚀 Готово! Файл index.html успешно собран.")
