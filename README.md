# 📄➡️📊 Word/PDF → Excel Veri Çıkarıcı

Groq AI kullanarak Word ve PDF dosyalarından doğal dil talimatıyla veri çıkarıp Excel tablosuna dönüştüren otomasyon aracı. Masaüstü arayüzü ve otomatik klasör izleme özelliğiyle çalışır.

## 🚀 Özellikler

- **Masaüstü GUI** — tkinter ile kolay kullanım, dosya seçici, canlı log
- **Klasör izleme** — klasöre dosya atılınca otomatik işler
- **Word ve PDF desteği** — .docx ve .pdf dosyalarını okur
- **Doğal dil talimatı** — "kişi adları ve telefonları" gibi serbest ifadeyle çalışır
- **Güzel Excel çıktısı** — renkli başlıklar, zebra satırlar, otomatik sütun genişliği
- **Hızlı** — Groq'un ücretsiz API'si ile saniyeler içinde sonuç

## 📦 Kurulum

```bash
pip install -r requirements.txt
```

## 🔑 API Anahtarı

[console.groq.com](https://console.groq.com) adresinden ücretsiz API anahtarı alın.

## 💻 Kullanım

```bash
python main.py
```

1. API anahtarını gir
2. İzlenecek klasörü seç
3. Ne çıkartılsın yaz (örn: "çalışan adları ve maaşları")
4. **İzlemeyi Başlat** tıkla
5. Klasöre .docx veya .pdf at → Excel otomatik oluşur! 🎉

## 🛠️ Nasıl Çalışır?

```
Klasöre dosya atılır (.docx / .pdf)
    ↓  Watchdog algılar
Metin çıkarılır
    ↓  Groq LLM ile JSON çıkarımı
Yapısal JSON
    ↓  openpyxl ile Excel yazma
Excel Dosyası ✅ (aynı klasöre kaydedilir)
```
