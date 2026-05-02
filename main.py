#!/usr/bin/env python3
"""
Word → Excel Veri Çıkarıcı
Groq AI kullanarak Word dosyasından istenen verileri çıkarır ve Excel'e yazar.
"""

import sys
import json
import re
import argparse
from pathlib import Path

from docx import Document
from groq import Groq
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def word_oku(dosya_yolu: str) -> str:
    """Word dosyasından tüm metni çıkarır."""
    doc = Document(dosya_yolu)
    satirlar = []

    for paragraf in doc.paragraphs:
        if paragraf.text.strip():
            satirlar.append(paragraf.text.strip())

    for tablo in doc.tables:
        for satir in tablo.rows:
            hucre_metinleri = [h.text.strip() for h in satir.cells if h.text.strip()]
            if hucre_metinleri:
                satirlar.append(" | ".join(hucre_metinleri))

    return "\n".join(satirlar)


def groq_ile_cikart(metin: str, istek: str, api_key: str) -> dict:
    """Groq API ile Word metninden istenen veriyi JSON olarak çıkarır."""
    client = Groq(api_key=api_key)

    sistem_mesaji = """Sen bir veri çıkarma asistanısın. Kullanıcının verdiği metinden istenen bilgileri çıkarıp SADECE geçerli bir JSON nesnesi döndürürsün.

JSON formatı şu şekilde olmalı:
{
  "basliklar": ["Sütun1", "Sütun2", ...],
  "satirlar": [
    ["değer1", "değer2", ...],
    ...
  ]
}

Kurallar:
- Sadece JSON döndür, başka hiçbir şey yazma
- Türkçe karakterleri koru
- Veri bulunamazsa boş liste döndür: {"basliklar": [], "satirlar": []}
- Her satır aynı sayıda sütun içermeli"""

    kullanici_mesaji = f"""Aşağıdaki Word belgesi metninden şu bilgileri çıkar: {istek}

METIN:
{metin[:8000]}"""

    yanit = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": sistem_mesaji},
            {"role": "user", "content": kullanici_mesaji}
        ],
        temperature=0.1,
        max_tokens=4000
    )

    ham_yanit = yanit.choices[0].message.content.strip()

    # JSON bloğunu temizle
    ham_yanit = re.sub(r'^```json\s*', '', ham_yanit)
    ham_yanit = re.sub(r'\s*```$', '', ham_yanit)

    return json.loads(ham_yanit)


def excel_olustur(veri: dict, cikti_yolu: str):
    """Çıkarılan veriyi güzel formatlanmış bir Excel dosyasına yazar."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Çıkarılan Veri"

    # Stil tanımları
    baslik_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    baslik_fill = PatternFill("solid", start_color="2E75B6")
    baslik_hizalama = Alignment(horizontal="center", vertical="center", wrap_text=True)

    veri_font = Font(name="Arial", size=10)
    cift_satir_fill = PatternFill("solid", start_color="D6E4F0")
    tek_satir_fill = PatternFill("solid", start_color="FFFFFF")
    veri_hizalama = Alignment(vertical="center", wrap_text=True)

    ince_sinir = Side(style="thin", color="BDD7EE")
    sinir = Border(left=ince_sinir, right=ince_sinir, top=ince_sinir, bottom=ince_sinir)

    basliklar = veri.get("basliklar", [])
    satirlar = veri.get("satirlar", [])

    if not basliklar:
        ws["A1"] = "Veri bulunamadı."
        wb.save(cikti_yolu)
        return

    # Başlık satırı
    for col_idx, baslik in enumerate(basliklar, start=1):
        hucre = ws.cell(row=1, column=col_idx, value=baslik)
        hucre.font = baslik_font
        hucre.fill = baslik_fill
        hucre.alignment = baslik_hizalama
        hucre.border = sinir

    ws.row_dimensions[1].height = 30

    # Veri satırları
    for row_idx, satir in enumerate(satirlar, start=2):
        fill = cift_satir_fill if row_idx % 2 == 0 else tek_satir_fill
        for col_idx, deger in enumerate(satir, start=1):
            hucre = ws.cell(row=row_idx, column=col_idx, value=deger)
            hucre.font = veri_font
            hucre.fill = fill
            hucre.alignment = veri_hizalama
            hucre.border = sinir
        ws.row_dimensions[row_idx].height = 20

    # Sütun genişliklerini otomatik ayarla
    for col_idx, baslik in enumerate(basliklar, start=1):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        max_uzunluk = len(str(baslik))
        for satir in satirlar:
            if col_idx - 1 < len(satir):
                max_uzunluk = max(max_uzunluk, len(str(satir[col_idx - 1])))
        ws.column_dimensions[col_letter].width = min(max_uzunluk + 4, 50)

    # Başlık satırını dondur
    ws.freeze_panes = "A2"

    # Özet bilgi
    ws_ozet = wb.create_sheet("Özet")
    ws_ozet["A1"] = "Toplam Satır:"
    ws_ozet["B1"] = len(satirlar)
    ws_ozet["A2"] = "Toplam Sütun:"
    ws_ozet["B2"] = len(basliklar)
    for hucre in ["A1", "A2"]:
        ws_ozet[hucre].font = Font(name="Arial", bold=True)

    wb.save(cikti_yolu)
    print(f"✅ Excel dosyası oluşturuldu: {cikti_yolu}")
    print(f"   📊 {len(satirlar)} satır, {len(basliklar)} sütun")


def main():
    parser = argparse.ArgumentParser(
        description="Word dosyasından AI ile veri çıkarıp Excel'e yazar",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Örnekler:
  python main.py rapor.docx "kişi adları ve telefon numaraları" -k GROQ_API_KEY
  python main.py fatura.docx "ürün adı, miktar, fiyat" -k GROQ_API_KEY -o fatura_data.xlsx
        """
    )
    parser.add_argument("word_dosyasi", help="Kaynak Word (.docx) dosyasının yolu")
    parser.add_argument("istek", help="Çıkartılacak verinin doğal dil açıklaması")
    parser.add_argument("-k", "--api-key", required=True, help="Groq API anahtarı")
    parser.add_argument("-o", "--cikti", help="Çıktı Excel dosyası adı (varsayılan: cikti.xlsx)")

    args = parser.parse_args()

    word_yolu = Path(args.word_dosyasi)
    if not word_yolu.exists():
        print(f"❌ Hata: '{word_yolu}' bulunamadı.")
        sys.exit(1)

    cikti_yolu = args.cikti or word_yolu.stem + "_cikti.xlsx"

    print(f"📄 Word dosyası okunuyor: {word_yolu}")
    metin = word_oku(str(word_yolu))
    print(f"   {len(metin)} karakter okundu.")

    print(f"🤖 Groq ile veri çıkarılıyor: '{args.istek}'")
    veri = groq_ile_cikart(metin, args.istek, args.api_key)

    satirlar = veri.get("satirlar", [])
    if not satirlar:
        print("⚠️  Belirtilen kriterlere uygun veri bulunamadı.")
    else:
        print(f"   {len(satirlar)} kayıt bulundu.")

    print(f"📊 Excel oluşturuluyor: {cikti_yolu}")
    excel_olustur(veri, cikti_yolu)


if __name__ == "__main__":
    main()
