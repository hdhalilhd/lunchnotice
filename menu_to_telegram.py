import os
import datetime as dt
import requests
from openpyxl import load_workbook

# === AYARLAR ===
CHAT_ID = int(os.getenv("TELEGRAM_CHAT_ID", "1677402217"))
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
EXCEL_PATH = os.getenv("EXCEL_PATH", "Subat_2026_Yemek_Listesi.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "Şubat 2026 Menü")
SEND_TOMORROW = os.getenv("SEND_TOMORROW", "true").lower() == "true"
# =================

MONTHS = {
    "Ocak": 1, "Şubat": 2, "Mart": 3, "Nisan": 4, "Mayıs": 5, "Haziran": 6,
    "Temmuz": 7, "Ağustos": 8, "Eylül": 9, "Ekim": 10, "Kasım": 11, "Aralık": 12
}

def parse_tr_date(s):
    parts = str(s).strip().split()
    if len(parts) != 3:
        return None
    try:
        return dt.date(int(parts[2]), MONTHS[parts[1]], int(parts[0]))
    except:
        return None

def get_menu(target_date):
    wb = load_workbook(EXCEL_PATH, data_only=True)
    ws = wb[SHEET_NAME]

    for row in ws.iter_rows(min_row=2, values_only=True):
        t, gun, corba, ana, yard, tatli, ekstra = row
        d = parse_tr_date(t)
        if d == target_date:
            label = "Yarın" if SEND_TOMORROW else "Bugün"
            return (
                f"📌 {label} ({t} - {gun}) Menü:\n"
                f"🍲 Çorba: {corba}\n"
                f"🍽️ Ana: {ana}\n"
                f"🥗 Yardımcı: {yard}\n"
                f"🍮 Tatlı/İçecek: {tatli}\n"
                f"✅ Ekstra: {ekstra}"
            )

    return f"❗ Menü bulunamadı: {target_date}"

def send_telegram(msg):
    if not TOKEN:
        raise RuntimeError("TELEGRAM_BOT_TOKEN boş. GitHub Secrets'a ekle.")
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    r = requests.post(url, json={"chat_id": CHAT_ID, "text": msg}, timeout=20)
    r.raise_for_status()

if __name__ == "__main__":
    today = dt.date.today()
    target = today + dt.timedelta(days=1) if SEND_TOMORROW else today
    send_telegram(get_menu(target))
