# Gümrük Beyanname Hazırlık Sistemi

CI (fatura) + PL (çeki listesi) → tarife bazında beyanname hazırlık Excel'i.
Otomatik tedarikçi/menşe/fatura-no tanıma, mutabakat, hata yakalama, öğrenme.

## Klasör yapısı
```
gumruk-sistemi/
├── scripts/beyanname_hazirla.py   ← ANA ARAÇ
├── scripts/arsiv_guncelle.py      ← arşiv referansı (nadiren)
├── veri/reference.pkl             ← geçmiş beyanname motoru
├── veri/ogrenilmis_eslesmeler.xlsx ← EN DEĞERLİ DOSYAN (kullandıkça büyür)
├── gelen/                         ← her iş kendi klasöründe
├── tamamlanan/                    ← onaylanan işler buraya taşınır
└── cikti/                         ← üretilen hazırlık Excel'leri
```

## KULLANIM — çok basit

### 1. İşi hazırla
`gelen/` içinde her iş için bir klasör aç, CI ve PL'yi içine koy:
```
gelen/GODSON_ocak/
├── CI_....xls
└── PL_....xls
```
Klasör adı önemli değil, sen anlarsın diye. CI/PL dosya adında "CI"/"invoice" ve
"PL"/"packing" geçsin ki sistem ayırt etsin (birleşik tek dosya da olur).

### 2. İşle (Cowork'e tek cümle)
> gelendeki dosyayı işle

Sistem otomatik: klasörü bulur, CI'dan tedarikçiyi/menşeyi/fatura-no'yu çıkarır,
eşleştirir, `cikti/`ya `[FİRMA]_[FATURANO]_HAZIRLIK.xlsx` üretir.
(Birden çok iş varsa: > gelen/GODSON_ocak klasörünü işle)

### 3. Kontrol et
`cikti/`daki Excel'i aç. **Kontrol sayfasına bak** — sarı "İNCELE" varsa Detay'da
o satırı gözden geçir (isim farkı / adet uyuşmazlık / GTİP boş / arşiv uyuşmazlık).

### 4. Onayla + öğret
Doğruysa Cowork'e:
> aynı işi --onayla ile çalıştır

Bu: öğrenme dosyasını günceller (sonraki aynı tedarikçi faturası otomatik gelir)
+ işi `tamamlanan/`a taşır (CI/PL/çıktı bir arada).

## Otomatik kurallar
- **Tedarikçi** = faturayı kesen firma (üstteki satıcı). Otomatik bulunur.
- **Menşe** = faturada yazan ülke; yoksa varsayılan 720 (Çin).
- **Dosya adı** = FİRMA_FATURANO (çakışmayı önler). Aynı ad varsa _2, _3 eklenir (üstüne YAZILMAZ).
- **GTİP** öncelik: (1) faturada varsa al+arşivle kontrol et (2) öğrenilmişse otomatik
  (3) arşivde benzer varsa öner (4) yoksa boş → sen gir.

## Renk kodları (Detay sayfası)
- YEŞİL: güvenilir (CI'dan / öğrenilmiş / arşiv tutuyor)
- SARI: dikkat / arşiv önerisi / sen onayla
- KIRMIZI: fark var / GTİP yok

## Elle geçmek istersen (opsiyonel)
```
python3 scripts/beyanname_hazirla.py --klasor gelen/X --tedarikci "FIRMA" --mense 720
```

## Notlar
- PDF fatura gelirse (nadir): önce Excel'e çevir, sonra klasöre koy.
- SKD/çok-modelli faturalar: sistem net olanları doldurur, belirsizleri işaretler; gerisi elle.
- Öğrenme dosyasını yedekle (OneDrive zaten yapıyor).
