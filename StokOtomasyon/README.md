# Stok Otomasyon (C# / WPF)

Excel tabanlı stok takip, sipariş ve hediyelik kayıt uygulaması. Verilerin tamamı
bir Excel dosyasında tutulur; uygulama bu dosyayı okur ve yazar, formüllere dokunmaz.

Depodaki `ornek_veri.xlsx` dosyası çalışan bir örnek şablondur — içindeki ürün
adları, kişi isimleri ve miktarların tamamı kurgusaldır.

## Sekmeler ve Özellikler

**📊 Stok Durumu**
Özet kartları (ürün sayısı, toplam kalan, kritik, tükendi), ürün arama,
"sadece kritik/tükenen" filtresi, renkli durum tablosu. Kritik/tükenen ürün varsa
sekme başlığında ⚠ uyarı rozeti görünür.

**🧾 Sipariş Girişi**
Kişi adı için otomatik tamamlama (önceki kişiler listeden seçilir), ürün arama kutusu,
tek tıkla düzenlenebilen adet hücreleri (çift tıklamaya gerek yok), canlı sepet özeti
(çeşit/adet toplamı anlık görünür), Temizle butonu. Stok yetersizse kaydetmeden önce uyarır.

**🕘 Sipariş Geçmişi**
Tüm siparişler tarihe göre sıralı listelenir. Filtreler: kişi (otomatik tamamlama),
tarih aralığı, neden. Satıra tıklayınca sağ panelde ürün kalemleri açılır.
Seçili filtreye göre kişi özeti (kaç sipariş, toplam adet) gösterilir.
**CSV dışa aktarma** (Türkçe Excel uyumlu) ve **sipariş silme** (adetler stoğa geri döner).

**🎁 Hediyelikler**
Sipariş girişine benzer: kişi (otomatik tamamlama), tarih ve ürün adetleri girilir;
hediye edilen adetler stoktan düşer. Aynı ekranda hediye geçmişi listelenir,
hatalı kayıtlar silinebilir (adetler stoğa geri döner). Excel'de ayrı
"Hediyelikler" sayfasına yazılır; Stok sayfasında "HEDİYE EDİLEN" sütununda görünür.

**➕ Stok Girişi**
Giriş formu + tüm giriş geçmişi bir arada. Hatalı girişler silinebilir.

**📈 Raporlar**
Tarih aralığı seçilebilir. Özet kartları (sipariş sayısı, toplam adet, en çok giden ürün,
en aktif kişi), en çok sipariş edilen 10 ürün için çubuk grafik, kişi bazında özet tablosu
(sipariş sayısı, toplam adet, son sipariş tarihi).

**🆕 Ürün Ekle**
Ürün hem Stok sayfasına satır hem Sipariş sayfasına sütun olarak eklenir.

Uygulama son kullanılan Excel dosyasını hatırlar, Excel formüllerine dokunmaz;
dosyayı Excel'de açtığınızda her şey otomatik yeniden hesaplanır. Excel'de elle
yapılan kayıtlarla tam uyumludur.

## Gereksinimler

- Windows 10/11
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) (veya Visual Studio 2022)

## Çalıştırma

```bash
cd StokOtomasyon
dotnet run
```

Kalıcı bir .exe üretmek için:

```bash
dotnet publish -c Release -r win-x64 --self-contained false
# Çıktı: bin/Release/net8.0-windows/win-x64/publish/StokOtomasyon.exe
```

İlk açılışta sağ üstteki **Excel Dosyası Seç…** butonuyla `ornek_veri.xlsx`
dosyasını (veya kendi çalışma dosyanızı) gösterin.

## Akıllı düzen tanıma

Uygulama satır/sütun numaralarını ezbere bilmez; her açılışta dosyanın içeriğine bakarak
düzeni kendisi keşfeder (`SablonHaritasi.cs`):

- Sayfalar **adlarına göre değil içeriklerine göre** tanınır: "VERİLEN KİŞİ AD SOYAD"
  başlığı olan sayfa Sipariş, "ÜRÜN ADI" + "BAŞLANGIÇ STOK" olan Stok, "TARİH" +
  "ÜRÜN ADI" + "ADET" olan Stok Giriş sayfasıdır. Sayfaları yeniden adlandırabilirsiniz.
- Başlık hücreleri metinden bulunur; başına satır eklemek, sütunların yerini
  değiştirmek, araya sütun eklemek uygulamayı bozmaz.
- Veri alanının sonu formüllerle hazırlanmış satırlardan tespit edilir; Stok Giriş
  kapasitesi, Stok sayfasındaki SUMIF formülünün okuduğu aralıktan otomatik öğrenilir.
- Düzen tanınamazsa hangi başlığın eksik olduğunu söyleyen açık bir hata gösterilir.

## Önemli notlar

- Kayıt sırasında Excel dosyası **Excel'de açık olmamalı** (dosya kilitlenir).
- Başlık metinlerini ("ÜRÜN ADI", "VERİLEN KİŞİ AD SOYAD", "TESLİM TARİHİ", "NEDENİ",
  "TOPLAM", "TARİH", "ADET") tamamen silmeyin/değiştirmeyin — uygulama düzeni bu
  metinlerden tanır.
- Sipariş/giriş silme işlemi Excel'de ilgili satırı temizler; formüller sayesinde
  stok otomatik geri hesaplanır.
