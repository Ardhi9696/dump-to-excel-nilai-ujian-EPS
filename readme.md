# 🔍 EPS-TOPIK Result Scraper 🇰🇷

Skrip ini digunakan untuk **mengambil data hasil ujian EPS-TOPIK** (Korea) secara otomatis dari situs resmi EPS menggunakan daftar nomor ujian dalam file Excel. Hasil scraping kemudian disimpan ke file Excel baru secara terstruktur.

---

## 📂 Fitur

- ✅ Scraping otomatis dari [https://www.eps.go.kr](https://www.eps.go.kr)
- 📥 Input dari file Excel (`no_ujian_5.xlsx`)
- 📊 Menyimpan hasil ke file `hasil_eps_topik.xlsx`
- 💡 Menampilkan hasil lengkap seperti:
  - Nama
  - Negara
  - Bidang
  - Nilai Listening, Reading
  - Total nilai & status kelulusan
  - Masa berlaku ujian

---

## 🧰 Kebutuhan

- Python 3.8+
- Google Chrome
- ChromeDriver (pastikan versi cocok dengan Chrome kamu)

### 🧪 Dependensi Python:

Install dengan pip:

```bash
pip install -r requirements.txt

⚠️ Catatan
Skrip ini bersifat scraping otomatis. Harap gunakan dengan bijak dan tidak berlebihan untuk menghormati server EPS.

Jika situs EPS lambat merespons, tambahkan time.sleep() di antara request.

Tidak mendukung CAPTCHA.