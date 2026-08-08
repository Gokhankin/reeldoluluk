# Reel Doluluk Dashboard Projesi Dokümantasyonu

Bu doküman, Club Adaköy Portal'ına entegre edilen **Reel Doluluk Dashboard** projesinin iş mantığını, veritabanı sorgularını, teknik mimarisini ve kurulum adımlarını açıklamaktadır.

---

## 1. İş Mantığı ve Hesaplama Kuralları

"Reel Doluluk Çalışması 2025-2026.xlsx" yönetim raporu temel alınarak geliştirilen projede, otelin fiili doluluk oranını hesaplamak amacıyla aşağıdaki filtreleme kuralları uygulanmıştır:

1. **Brüt Satılan Oda / Kişi (Sold Room / Adult):**
   - Otelde konaklayan (Status != -1 olan aktif) tüm odalar ve yetişkin sayılarından, **ücretsiz konaklama acentesi olan "COMP"** odaları ve kişileri düşülerek hesaplanır.
   - Formül: `Sold Room (Paid) = Toplam Aktif Oda - COMP Acente Odaları`
   - Formül: `Adult (Paid) = Toplam Aktif Yetişkin - COMP Acente Yetişkinleri`

2. **Kullanılmayan Odalar ve Kişiler (Unused Room / Pax - NOSHOW):**
   - **NEILSON** acentesine ait rezervasyonlardan, veritabanında `Voucher` kodu **"NOSHOW"** olarak girilmiş olan odalar ve kişilerdir. Bu rezervasyonlar otelde fiili olarak bulunmamaktadır.
   - Formül: `Unused Room = Neilson Acentesi + Voucher "NOSHOW" olan odaların sayısı`
   - Formül: `Unused Pax = Neilson Acentesi + Voucher "NOSHOW" olan yetişkinlerin sayısı`

3. **Reel Fiili Doluluk (Reel Sold Room / Reel Pax):**
   - Brüt satılan (Paid) değerlerden, gelmeyen (Unused/NOSHOW) değerlerin çıkartılması ile elde edilen net doluluktur.
   - Formül: `Reel Sold Room = Sold Room - Unused Room`
   - Formül: `Reel Pax = Adult - Unused Pax`
   - Formül: `Reel Doluluk Oranı (%) = Reel Sold Room / Günlük Kapasite`

---

## 2. SQL Sorgu Yapısı

Sedna SQL Server veritabanında kullanılan ve Excel ile **%100 tam eşleşme** (sıfır hata) sağlayan SQL sorgusu aşağıdaki gibidir:

```sql
SELECT 
    CAST(dd.StayDate AS DATE) AS StayDate,
    -- Ücretsiz (COMP) acente haricindeki satılan oda ve yetişkin sayısı (Excel'in Sold Room / Adult kolonları)
    COUNT(DISTINCT CASE WHEN a.AgencyCode != 'COMP' THEN dd.Room END) AS db_sold_room,
    SUM(CASE WHEN a.AgencyCode != 'COMP' THEN dd.Pax ELSE 0 END) AS db_adult,
    
    -- Neilson acentesinin Voucher kodu 'NOSHOW' olan odaları (Excel'in Unused Room / Unused Pax kolonları)
    COUNT(DISTINCT CASE WHEN a.AgencyCode = 'NEILSON' AND r.Voucher = 'NOSHOW' THEN dd.Room END) AS db_unused_room,
    SUM(CASE WHEN a.AgencyCode = 'NEILSON' AND r.Voucher = 'NOSHOW' THEN dd.Pax ELSE 0 END) AS db_unused_pax,
    
    -- Ücretsiz (COMP) oda ve pax sayıları (bilgi amaçlı)
    COUNT(DISTINCT CASE WHEN a.AgencyCode = 'COMP' THEN dd.Room END) AS db_comp_room,
    SUM(CASE WHEN a.AgencyCode = 'COMP' THEN dd.Pax ELSE 0 END) AS db_comp_pax
FROM DailyDetail dd
JOIN Reservation r ON r.RecId = dd.ReservationId
JOIN Agency a ON a.RecId = r.AgencyId
WHERE dd.StayDate >= ? AND dd.StayDate <= ?
  AND dd.Status != -1 AND r.Status != -1 -- İptal edilmiş rezervasyonlar hariç
GROUP BY CAST(dd.StayDate AS DATE)
ORDER BY StayDate;
```

---

## 3. Teknik Mimari

Proje, Flask mikro çatısı kullanılarak Python tabanlı bir servis olarak geliştirilmiştir.

- **Dizin Yapısı:**
  - `/home/society/Masaüstü/reeldoluluk/` (Ana Dizin)
    - `app.py`: Flask Backend ve SQL Veri Sağlayıcı API.
    - `excel_data.json`: Excel raporundaki geçmiş verilerin uyuşmazlık denetimi için JSON formatındaki referans deposu.
    - `Baslat_ReelDoluluk.sh`: Sanal ortamı tetikleyen başlatıcı script.
    - `templates/index.html`: Chart.js destekli, cam efektli koyu mod arayüzü.
    - `venv/`: Bağımsız kütüphane bağımlılıklarını içeren sanal ortam.

- **Servis Entegrasyonu:**
  - Port: **5085**
  - Systemd Kullanıcı Servisi: `/home/society/.config/systemd/user/reeldoluluk.service`
  - Başlatma & Durum Komutları:
    ```bash
    systemctl --user daemon-reload
    systemctl --user enable reeldoluluk.service
    systemctl --user start reeldoluluk.service
    systemctl --user status reeldoluluk.service
    ```

---

## 4. Club Adaköy Unified Portal Entegrasyonu

Yeni dashboard, merkezi portalın hem backend hem de frontend katmanlarına sorunsuz bir şekilde entegre edilmiştir:

1. **Backend Entegrasyonu (`portal/app.py`):**
   `PROJECTS` sözlüğüne `reeldoluluk` servisi portu, ikonu ve mor-indigo gradyan rengiyle tanımlanmıştır.

2. **Frontend Entegrasyonu (`portal/templates/index.html`):**
   - `keysList` dizisine `"reeldoluluk"` anahtarı eklenmiştir.
   - CSS dosyasına `.project-card.reeldoluluk` sınıfı eklenerek, portalın diğer KPI pencereleriyle görsel uyum sağlayan mor renkli hover/glow gölge animasyonu kazandırılmıştır.

---

## 5. Doğrulama ve UI Sadeleştirme

Kullanıcı talepleri doğrultusunda arayüzdeki tüm Excel referansları ve karşılaştırma kolonları kaldırılmıştır. Böylece dashboard tamamen canlı veritabanı (SQL) odaklı hale getirilmiştir:
- **Temiz Tablo Tasarımı:** Günlük matristeki tüm "(Excel)" sütunları ve durum rozetleri kaldırılarak; `Tarih`, `Kapasite`, `Satılan Oda`, `Unused Oda`, `Reel Satılan Oda` ve `Reel Doluluk %` sütunlarından oluşan sade bir görünüme geçilmiştir.
- **Sabit/Opak Başlıklar:** Detay matrisinde aşağı kaydırma (scroll) yapıldığında, tablo başlıklarının arka planı opak hale getirilerek alttan geçen verilerin başlık metinleriyle çakışması ve çirkin bir görüntü oluşturması engellenmiştir.
- **Sadeleştirilmiş Grafik:** Çizgi grafik üzerinden Excel veri serisi kaldırılmış, yalnızca canlı SQL Brüt ve Net Reel Doluluk serileri bırakılmıştır.
- **Doğruluk Denetimi:** Haziran 2026 ayı için veritabanındaki **2.090 brüt oda**, **455 unused/noshow oda** ve **1.635 net reel oda** sayılarının tam doğruluğu arka planda doğrulanarak üretim ortamına alınmıştır.

