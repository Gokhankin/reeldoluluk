# Reel Doluluk Analiz Raporu Projesi Dokümantasyonu (reel.md)

Bu doküman, Club Adaköy Portal'ına entegre edilen ve `192.168.0.128:5085` adresinde canlı hizmet veren **Reel Doluluk Analiz Raporu** projesinin iş mantığını, veritabanı sorgularını, API uç noktalarını, Sezonluk grafik & kümülatif tablo mimarisini ve güncel arayüz bileşenlerini açıklamaktadır.

---

## 1. İş Mantığı ve Hesaplama Kuralları

"Reel Doluluk Çalışması 2025-2026 (2).xlsx" yönetim raporu temel alınarak geliştirilen projede, otelin fiili net doluluk oranını hesaplamak amacıyla aşağıdaki filtreleme kuralları uygulanmıştır:

1. **Brüt Satılan Oda / Kişi (Sold Room / Adult):**
   - Otelde konaklayan (Status != -1 olan aktif) tüm odalar ve yetişkin sayılarından, **ücretsiz konaklama acentesi olan "COMP"** odaları ve kişileri düşülerek hesaplanır.
   - Formül: `Sold Room (Paid) = Toplam Aktif Oda - COMP Acente Odaları`
   - Formül: `Adult (Paid) = Toplam Aktif Yetişkin - COMP Acente Yetişkinleri`

2. **Kullanılmayan Odalar ve Kişiler (Unused Room / Pax - NOSHOW):**
   - **NEILSON** acentesine ait rezervasyonlardan, veritabanında `Voucher` kodu **"NOSHOW"** (boşluk duyarsız / space-insensitive) olarak girilmiş olan odalar ve kişilardır. Bu rezervasyonlar otelde fiili olarak bulunmamaktadır.
   - Formül: `Unused Room = Neilson Acentesi + Voucher "NOSHOW" olan odaların sayısı`
   - Formül: `Unused Pax = Neilson Acentesi + Voucher "NOSHOW" olan yetişkinlerin sayısı`

3. **Reel Fiili Doluluk (Reel Sold Room / Reel Pax):**
   - Brüt satılan (Paid) değerlerden, gelmeyen (Unused/NOSHOW) değerlerin çıkartılması ile elde edilen fiili net doluluktur.
   - Formül: `Reel Sold Room = Sold Room - Unused Room`
   - Formül: `Reel Pax = Adult - Unused Pax`
   - Formül: `Reel Doluluk Oranı (%) = Reel Sold Room / Günlük Kapasite`

---

## 2. API ve Veri Mimarisi

Sistem Flask mikro çatısı ile çalışmaktadır ve 2 temel API uç noktası sunar:

### A. `/api/data?year=YYYY&month=MM` (Aylık Günlük Detay API)
- Seçili yıl ve ay için gün bazında Kapasite, Brüt Satılan Oda, Unused (Noshow) Oda, Reel Satılan Oda, Yetişkin ve Pax sayılarını döndürür.

### B. `/api/seasonal_data?year=YYYY` (Sezonluk Kümülatif API)
- Tüm sezon boyunca (Mayıs - Ekim) aylık bazda aggregate edilmiş veriyi döndürür:
  - **2026 Sezonu:** Sedna SQL Server veritabanından `CheckinDate` ve `CheckOutDate` tarih aralıklarıyla dinamik sorgulanır.
  - **2025 Sezonu:** Geçmiş yıl Excel yönetim raporu ile %100 birebir uyumlu statik veri dizisi ile sunulur.

---

## 3. SQL Sorgu Yapısı (Sedna SQL Server)

Sedna SQL Server veritabanında kullanılan ve Excel raporları ile **%100 tam eşleşme** sağlayan SQL sorgusu:

```sql
SELECT 
    CAST(dd.StayDate AS DATE) AS StayDate,
    -- Ücretsiz (COMP) acente haricindeki satılan oda ve yetişkin sayısı
    COUNT(DISTINCT CASE WHEN a.AgencyCode != 'COMP' THEN dd.Room END) AS db_sold_room,
    SUM(CASE WHEN a.AgencyCode != 'COMP' THEN dd.Pax ELSE 0 END) AS db_adult,
    
    -- Neilson acentesinin Voucher kodu 'NOSHOW' olan odaları
    COUNT(DISTINCT CASE WHEN a.AgencyCode = 'NEILSON' AND UPPER(REPLACE(ISNULL(r.Voucher, ''), ' ', '')) = 'NOSHOW' THEN dd.Room END) AS db_unused_room,
    SUM(CASE WHEN a.AgencyCode = 'NEILSON' AND UPPER(REPLACE(ISNULL(r.Voucher, ''), ' ', '')) = 'NOSHOW' THEN dd.Pax ELSE 0 END) AS db_unused_pax
FROM DailyDetail dd
JOIN Reservation r ON r.RecId = dd.ReservationId
JOIN Agency a ON a.RecId = r.AgencyId
WHERE dd.StayDate >= @StartDate AND dd.StayDate <= @EndDate
  AND dd.Status != -1 AND r.Status != -1
GROUP BY CAST(dd.StayDate AS DATE)
ORDER BY StayDate;
```

---

## 4. UI/UX ve Grafik Yenilikleri

1. **Dinamik Aktif Ay Algılama:**
   * Sayfa açıldığında sistem otomatik olarak içinde bulunulan aktif ayı (örn. Eylül) seçili getirir.

2. **KPI Kartları & Tipografi:**
   * Ana KPI kartlarında Oda ve Pax sayıları eşit büyük font boyutunda yan yana gösterilir:
     * `2.549 Oda / 4.702 Pax` (Brüt)
     * `213 Oda / 426 Pax` (Noshow/Unused)
     * `2.336 Oda / 4.276 Pax` (Reel Net)
   * Alt kısımlarda derece/doluluk oranlarını canlı dolduran renkli **animasyonlu ilerleme çubukları (Progress Bars)** yer alır.

3. **Tam Genişlikte Çizgi Grafik (Full Width):**
   * Bilgilendirme paneli kaldırılarak *"Seçili Ay Günlük Doluluk Grafiği (%)"* ekranı %100 genişlikle kaplayacak şekilde tasarlanmıştır.

4. **Sezonluk Sütun Grafiği (Bar Chart):**
   * Unused Room (Mavi) ve Reel Sold Room (Turuncu) sayıları aylık sütunlarda gösterilir.
   * `Chart.js DataLabels` eklentisi ile her sütunun üzerinde net değerler etiketlenmiştir.

5. **Sezonluk Kümülatif Özet Tablosu:**
   * **FORECAST**, **UNUSED OLARAK İŞLENEN** ve **REEL DOLULUK** başlık gruplarıyla aylık dağılım ve en altta **TOPLAM (TOTAL)** satırı sunulur.

6. **Mobil Uyumluluk (Responsive):**
   * Tüm tablolar ve grafikler mobile özel dokunmatik kaydırma (`overflow-x: auto`) desteklidir.

---

## 5. Canlı Sunucu ve Kurulum Bilgileri

- **Sunucu IP:** `192.168.0.128` (User: `society`)
- **Proje Dizini:** `/home/society/Masaüstü/reeldoluluk/`
- **Canlı Port:** `5085`
- **Canlı Uygulama Adresi:** `http://192.168.0.128:5085/`
- **Systemd Servis Yolu:** `/home/society/.config/systemd/user/reeldoluluk.service`
- **Servis Yönetim Komutları:**
  ```bash
  systemctl --user restart reeldoluluk.service
  systemctl --user status reeldoluluk.service
  journalctl --user -u reeldoluluk.service -f
  ```
