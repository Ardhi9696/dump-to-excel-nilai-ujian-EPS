import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from tkinter import Tk, filedialog, messagebox
from playsound import playsound

# === UI Setup ===
root = Tk()
root.withdraw()  # Sembunyikan jendela utama

# === Pilih file Excel input ===
input_file = filedialog.askopenfilename(
    title="Pilih File Excel Nomor Ujian",
    filetypes=[("Excel Files", "*.xlsx *.xls")],
)
if not input_file:
    print("❌ Tidak ada file dipilih. Keluar.")
    exit()

# === Setup Headless Chrome ===
options = Options()
options.add_argument("--headless")
options.add_argument("--window-size=1920,1080")
driver = webdriver.Chrome(options=options)

# === Baca file Excel input ===
df_input = pd.read_excel(input_file)
nomor_list = df_input.iloc[:, 0].astype(str).tolist()

hasil_data = []

for nomor_ujian in nomor_list:
    print(f"\n🔍 Memproses nomor ujian: {nomor_ujian}")
    try:
        driver.get("https://www.eps.go.kr/eo/VisaFndRM.eo?langType=in")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "sKorTestNo"))
        )

        input_box = driver.find_element(By.ID, "sKorTestNo")
        input_box.clear()
        input_box.send_keys(nomor_ujian)

        print("➡️ Klik tombol View...")
        submit_button = driver.find_element(
            By.XPATH, "//button[contains(text(),'View')]"
        )
        submit_button.click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "tbl_typeA"))
        )
        soup = BeautifulSoup(driver.page_source, "html.parser")
        table = soup.select_one(".tbl_typeA")

        rows = table.find_all("tr")
        cells = [
            cell.get_text(strip=True) for row in rows for cell in row.find_all("td")
        ]

        if len(cells) >= 12:
            hasil = {
                "Nomor Ujian": nomor_ujian,
                "Nama": cells[5],
                "Negara": cells[1],
                "Sektor": cells[2],
                "Tanggal Ujian": cells[3],
                "Listening": cells[6],
                "Reading": cells[7],
                "Total Nilai": cells[8],
                "KKM": cells[9],
                "Lulus": cells[10],
                "Masa Berlaku": cells[11],
            }
            print(f"✅ Ditemukan: {hasil['Nama']} - {hasil['Total Nilai']}")
        else:
            hasil = {
                "Nomor Ujian": nomor_ujian,
                "Nama": "-",
                "Negara": "-",
                "Sektor": "-",
                "Tanggal Ujian": "-",
                "Listening": "-",
                "Reading": "-",
                "Total Nilai": "-",
                "KKM": "-",
                "Lulus": "-",
                "Masa Berlaku": "-",
            }
            print("❌ Data tidak lengkap / tidak ditemukan.")
    except Exception as e:
        hasil = {
            "Nomor Ujian": nomor_ujian,
            "Nama": "ERROR",
            "Negara": "",
            "Sektor": "",
            "Tanggal Ujian": "",
            "Listening": "",
            "Reading": "",
            "Total Nilai": "",
            "KKM": "",
            "Lulus": "",
            "Masa Berlaku": "",
        }
        print(f"❌ Gagal memproses: {e}")
    hasil_data.append(hasil)

driver.quit()

# === Simpan hasil dengan Save-As Dialog ===
output_file = filedialog.asksaveasfilename(
    title="Simpan Hasil Scraping",
    defaultextension=".xlsx",
    filetypes=[("Excel Files", "*.xlsx *.xls")],
)

if output_file:
    kolom_urutan = [
        "Nomor Ujian",
        "Nama",
        "Negara",
        "Sektor",
        "Tanggal Ujian",
        "Listening",
        "Reading",
        "Total Nilai",
        "KKM",
        "Lulus",
        "Masa Berlaku",
    ]
    df = pd.DataFrame(hasil_data)
    df.to_excel(output_file, index=False, columns=kolom_urutan)
    print(f"\n✅ Selesai! Hasil disimpan ke '{output_file}'")

    # 🔔 Notifikasi suara
    try:
        playsound("notif.mp3")  # Ganti dengan file suara kamu
    except Exception:
        os.system("afplay /System/Library/Sounds/Glass.aiff")  # macOS fallback
        # os.system('start /min mplay32 /play /close "notif.wav"')  # Windows fallback

    # ✅ Pop-up konfirmasi
    messagebox.showinfo(
        "Selesai", "🎉 Scraping selesai!\nFile hasil berhasil disimpan."
    )
else:
    print("❌ Simpan hasil dibatalkan.")


# import pandas as pd
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from bs4 import BeautifulSoup
# import time

# # Konfigurasi headless
# options = Options()
# options.add_argument("--headless")
# options.add_argument("--window-size=1920,1080")
# driver = webdriver.Chrome(options=options)

# # Baca file Excel
# df_input = pd.read_excel("sources/ubt_special.xlsx")
# nomor_list = df_input.iloc[:, 0].astype(str).tolist()  # Ambil kolom pertama

# # Simpan hasil scraping
# hasil_data = []

# for nomor_ujian in nomor_list:
#     print(f"\n🔍 Memproses nomor ujian: {nomor_ujian}")
#     try:
#         driver.get("https://www.eps.go.kr/eo/VisaFndRM.eo?langType=in")
#         WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.ID, "sKorTestNo"))
#         )

#         input_box = driver.find_element(By.ID, "sKorTestNo")
#         input_box.clear()
#         input_box.send_keys(nomor_ujian)

#         print("➡️  Klik tombol View...")
#         submit_button = driver.find_element(
#             By.XPATH, "//button[contains(text(),'View')]"
#         )
#         submit_button.click()

#         WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.CLASS_NAME, "tbl_typeA"))
#         )
#         soup = BeautifulSoup(driver.page_source, "html.parser")
#         table = soup.select_one(".tbl_typeA")

#         rows = table.find_all("tr")
#         cells = [
#             cell.get_text(strip=True) for row in rows for cell in row.find_all("td")
#         ]

#         if len(cells) >= 12:
#             hasil = {
#                 "Nomor Ujian": nomor_ujian,
#                 "Nama": cells[5],
#                 "Negara": cells[1],
#                 "Bidang": cells[2],
#                 "Tanggal Ujian": cells[3],
#                 "Mendengar": cells[6],
#                 "Bacaan": cells[7],
#                 "Total Nilai": cells[8],
#                 "Nilai Lulus": cells[9],
#                 "Lulus": cells[10],
#                 "Masa Berlaku": cells[11],
#             }
#             print(f"✅ Ditemukan: {hasil['Nama']} - {hasil['Total Nilai']}")
#         else:
#             hasil = {
#                 "Nomor Ujian": nomor_ujian,
#                 "Nama": "-",
#                 "Negara": "-",
#                 "Bidang": "-",
#                 "Tanggal Ujian": "-",
#                 "Listening": "-",
#                 "Reading": "-",
#                 "Total Nilai": "-",
#                 "KKM": "-",
#                 "Lulus": "-",
#                 "Masa Berlaku": "-",
#             }
#             print("❌ Data tidak lengkap / tidak ditemukan.")
#     except Exception as e:
#         hasil = {
#             "Nomor Ujian": nomor_ujian,
#             "Nama": "ERROR",
#             "Negara": "",
#             "Bidang": "",
#             "Tanggal Ujian": "",
#             "Mendengar": "",
#             "Bacaan": "",
#             "Total Nilai": "",
#             "Nilai Lulus": "",
#             "Lulus": "",
#             "Masa Berlaku": "",
#         }
#         print(f"❌ Gagal memproses: {e}")
#     hasil_data.append(hasil)

# # Simpan hasil akhir
# pd.DataFrame(hasil_data).to_excel("results/hasil_eps_topik_special.xlsx", index=False)
# print("\n✅ Selesai! Hasil disimpan ke 'hasil_eps_topik.xlsx'")
# driver.quit()
