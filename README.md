# Mizandan Gelir Tablosu ve Bilanço

Finansal analiz raporu şablonundaki **Gelir Tablosu** ve **Bilanço** sayfalarını,
verilen mizanlardan otomatik dolduran web uygulaması.

## Kullanım

1. Şablonu Excel'de aç
2. En sona iki sayfa ekle:
   - `mizan1` → geçen yılın mizanı
   - `mizan2` → bu yılın mizanı
3. Dosyayı uygulamaya yükle, sonucu indir

Sayfa adları **tam olarak** `mizan1` ve `mizan2` olmalıdır (küçük harf, boşluksuz).
`mizan2` yoksa program çalışmaz. `mizan1` yoksa önceki dönem sütunu boş bırakılır.

## Ne yapar

- 7/A maliyet hesaplarındaki yansıtılmamış kalanı 6'lı hesaplara aktarır
  (`Kalan = 7'li borç bakiyesi − yansıtma alacak bakiyesi`)
- Dönem kâr/zararını hesaplayıp 590 / 591'e yazar
- Aktif = Pasif denkliğini kontrol eder
- Mizan sayfalarını siler

## Ne yapmaz

- **Vergi karşılığı hesaplamaz** (geçici vergi dönemi tablolarıdır)
- Üretim/stok maliyeti hesaplamaz
- Şablonun biçimlendirmesine dokunmaz

## Yerel çalıştırma

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Dosyalar

| Dosya | Görevi |
|---|---|
| `app.py` | Streamlit arayüzü |
| `isleyici.py` | Ana işlem akışı |
| `mali_mantik.py` | Mizan okuma, 7/A aktarımı, hesaplamalar |
| `xlsx_cerrahi.py` | Biçimlendirmeyi bozmadan Excel'e yazma |
