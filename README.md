# 📄➡️📊 Word → Excel Veri Çıkarıcı

Groq AI kullanarak Word dosyasından doğal dil talimatıyla veri çıkarıp Excel tablosuna dönüştüren otomasyon aracı.

## 🚀 Özellikler

- **Karışık belgeler desteklenir** – rapor, fatura, liste, tablo, ne olursa
- **Doğal dil talimatı** – "kişi adları ve telefonları" gibi serbest ifadeyle çalışır
- **Güzel Excel çıktısı** – renkli başlıklar, zebra satırlar, otomatik sütun genişliği
- **Özet sayfası** – kaç satır/sütun bulunduğunu gösterir
- **Hızlı** – Groq'un ücretsiz API'si ile saniyeler içinde sonuç

## 📦 Kurulum

```bash
pip install -r requirements.txt
```

## 🔑 API Anahtarı

[console.groq.com](https://console.groq.com) adresinden ücretsiz API anahtarı alın.

## 💻 Kullanım

```bash
python main.py <word_dosyası> "<ne çıkartılsın>" -k <GROQ_API_KEY>
```

### Örnekler

```bash
# Çalışan listesinden ad ve departman çıkar
python main.py calisanlar.docx "çalışan adı ve departmanı" -k gsk_xxxx

# Toplantı notlarından karar ve sorumlu kişileri çıkar
python main.py toplanti_notu.docx "kararlar ve sorumlu kişiler" -k gsk_xxxx

# Faturadan ürün, miktar ve fiyat çıkar, özel isimle kaydet
python main.py fatura.docx "ürün adı, adet, birim fiyat" -k gsk_xxxx -o fatura_analiz.xlsx

# Araştırma raporundan bulgular çıkar
python main.py rapor.docx "bulgular ve öneriler" -k gsk_xxxx
```

### Parametreler

| Parametre | Açıklama |
|-----------|----------|
| `word_dosyasi` | Kaynak `.docx` dosyasının yolu |
| `istek` | Çıkartılacak verinin açıklaması (doğal dil) |
| `-k / --api-key` | Groq API anahtarı |
| `-o / --cikti` | Çıktı Excel dosya adı (varsayılan: `<dosyaadı>_cikti.xlsx`) |

## 📁 Çıktı

- **Veri sayfası**: Renkli başlıklı, zebra çizgili tablo
- **Özet sayfası**: Toplam satır ve sütun sayısı
- Başlık satırı dondurulmuş (scroll yaparken görünür kalır)

## 🛠️ Nasıl Çalışır?

```
Word Dosyası
    ↓  python-docx ile metin ve tablo okuma
Ham Metin
    ↓  Groq LLM (llama-3.3-70b) ile JSON çıkarımı
Yapısal JSON
    ↓  openpyxl ile Excel yazma
Excel Dosyası ✅
```

## 📝 Notlar

- Groq'un ücretsiz katmanında dakikada ~30 istek limiti vardır
- Çok büyük belgeler için metin 8000 karakterle sınırlandırılır
- Türkçe karakterler tam desteklenir
