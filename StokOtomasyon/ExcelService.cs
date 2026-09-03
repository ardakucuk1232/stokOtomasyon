using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace StokOtomasyon;

public class ExcelService
{
    public string DosyaYolu { get; }

    public ExcelService(string dosyaYolu)
    {
        DosyaYolu = dosyaYolu;
        if (!File.Exists(dosyaYolu))
            throw new FileNotFoundException("Excel dosyası bulunamadı: " + dosyaYolu);

        using var wb = new XLWorkbook(dosyaYolu);
        SablonHaritasi.Cikar(wb);
    }

    public List<StokUrun> StokOku()
    {
        using var wb = new XLWorkbook(DosyaYolu);
        var h = SablonHaritasi.Cikar(wb);

        var urunler = new List<StokUrun>();
        var indeks = new Dictionary<string, StokUrun>(StringComparer.OrdinalIgnoreCase);

        for (int r = h.StokIlkSatir; r <= h.StokSonSatir; r++)
        {
            string ad = h.Stok.Cell(r, h.StokUrunKol).GetString().Trim();
            if (ad.Length == 0) continue;
            var u = new StokUrun { Ad = ad, Baslangic = HucreSayi(h.Stok.Cell(r, h.StokBaslangicKol)) };
            urunler.Add(u);
            indeks[ad] = u;
        }

        for (int r = h.GirIlkSatir; r <= h.GirSonSatir; r++)
        {
            string ad = h.Giris.Cell(r, h.GirKolUrun).GetString().Trim();
            if (ad.Length == 0) continue;
            if (indeks.TryGetValue(ad, out var u))
                u.Giris += HucreSayi(h.Giris.Cell(r, h.GirKolAdet));
        }

        var kolonUrunu = new Dictionary<int, StokUrun>();
        for (int c = h.SipUrunIlkKol; c <= h.SipUrunSonKol; c++)
        {
            string baslik = h.Siparis.Cell(h.SipBaslikSatir, c).GetString().Trim();
            if (baslik.Length > 0 && indeks.TryGetValue(baslik, out var u))
                kolonUrunu[c] = u;
        }
        for (int r = h.SipIlkSatir; r <= h.SipSonSatir; r++)
        {
            if (h.Siparis.Cell(r, h.SipKolAd).GetString().Trim().Length == 0) continue;
            foreach (var (c, u) in kolonUrunu)
                u.Siparis += HucreSayi(h.Siparis.Cell(r, c));
        }

        var hedKolonUrunu = new Dictionary<int, StokUrun>();
        for (int c = h.HedUrunIlkKol; c <= h.HedUrunSonKol; c++)
        {
            string baslik = h.Hediye.Cell(h.HedBaslikSatir, c).GetString().Trim();
            if (baslik.Length > 0 && indeks.TryGetValue(baslik, out var u))
                hedKolonUrunu[c] = u;
        }
        for (int r = h.HedIlkSatir; r <= h.HedSonSatir; r++)
        {
            if (h.Hediye.Cell(r, h.HedKolAd).GetString().Trim().Length == 0) continue;
            foreach (var (c, u) in hedKolonUrunu)
                u.Hediye += HucreSayi(h.Hediye.Cell(r, c));
        }

        return urunler;
    }

    public List<SiparisKaydi> HediyeleriOku()
    {
        using var wb = new XLWorkbook(DosyaYolu);
        var h = SablonHaritasi.Cikar(wb);

        var kolonAdi = new Dictionary<int, string>();
        for (int c = h.HedUrunIlkKol; c <= h.HedUrunSonKol; c++)
        {
            string baslik = h.Hediye.Cell(h.HedBaslikSatir, c).GetString().Trim();
            if (baslik.Length > 0) kolonAdi[c] = baslik;
        }

        var liste = new List<SiparisKaydi>();
        for (int r = h.HedIlkSatir; r <= h.HedSonSatir; r++)
        {
            string kisi = h.Hediye.Cell(r, h.HedKolAd).GetString().Trim();
            if (kisi.Length == 0) continue;

            var kayit = new SiparisKaydi
            {
                SatirNo = r,
                Kisi = kisi,
                Tarih = HucreTarih(h.Hediye.Cell(r, h.HedKolTarih)),
                Neden = "HEDİYE"
            };
            foreach (var (c, ad) in kolonAdi)
            {
                int adet = HucreSayi(h.Hediye.Cell(r, c));
                if (adet != 0) kayit.Kalemler.Add(new KalemBilgi { Urun = ad, Adet = adet });
            }
            liste.Add(kayit);
        }
        return liste;
    }

    public List<SiparisKaydi> SiparisleriOku()
    {
        using var wb = new XLWorkbook(DosyaYolu);
        var h = SablonHaritasi.Cikar(wb);

        var kolonAdi = new Dictionary<int, string>();
        for (int c = h.SipUrunIlkKol; c <= h.SipUrunSonKol; c++)
        {
            string baslik = h.Siparis.Cell(h.SipBaslikSatir, c).GetString().Trim();
            if (baslik.Length > 0) kolonAdi[c] = baslik;
        }

        var liste = new List<SiparisKaydi>();
        for (int r = h.SipIlkSatir; r <= h.SipSonSatir; r++)
        {
            string kisi = h.Siparis.Cell(r, h.SipKolAd).GetString().Trim();
            if (kisi.Length == 0) continue;

            var kayit = new SiparisKaydi
            {
                SatirNo = r,
                Kisi = kisi,
                Tarih = HucreTarih(h.Siparis.Cell(r, h.SipKolTarih)),
                Neden = h.Siparis.Cell(r, h.SipKolNeden).GetString().Trim()
            };
            foreach (var (c, ad) in kolonAdi)
            {
                int adet = HucreSayi(h.Siparis.Cell(r, c));
                if (adet != 0) kayit.Kalemler.Add(new KalemBilgi { Urun = ad, Adet = adet });
            }
            liste.Add(kayit);
        }
        return liste;
    }

    public List<GirisKaydi> GirisleriOku()
    {
        using var wb = new XLWorkbook(DosyaYolu);
        var h = SablonHaritasi.Cikar(wb);

        var liste = new List<GirisKaydi>();
        for (int r = h.GirIlkSatir; r <= h.GirSonSatir; r++)
        {
            string urun = h.Giris.Cell(r, h.GirKolUrun).GetString().Trim();
            if (urun.Length == 0) continue;
            liste.Add(new GirisKaydi
            {
                SatirNo = r,
                Tarih = HucreTarih(h.Giris.Cell(r, h.GirKolTarih)),
                Urun = urun,
                Adet = HucreSayi(h.Giris.Cell(r, h.GirKolAdet)),
                Aciklama = h.Giris.Cell(r, h.GirKolAciklama).GetString().Trim()
            });
        }
        return liste;
    }

    public void UrunEkle(string ad, int baslangicStok)
    {
        ad = ad.Trim();
        if (ad.Length == 0) throw new ArgumentException("Ürün adı boş olamaz.");

        using var wb = new XLWorkbook(DosyaYolu);
        var h = SablonHaritasi.Cikar(wb);

        for (int r = h.StokIlkSatir; r <= h.StokSonSatir; r++)
            if (string.Equals(h.Stok.Cell(r, h.StokUrunKol).GetString().Trim(), ad,
                              StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"\"{ad}\" zaten stok listesinde var.");

        int bosSatir = -1;
        for (int r = h.StokIlkSatir; r <= h.StokSonSatir; r++)
            if (h.Stok.Cell(r, h.StokUrunKol).GetString().Trim().Length == 0) { bosSatir = r; break; }
        if (bosSatir < 0)
            throw new InvalidOperationException(
                "Stok sayfasında formüllerle hazırlanmış boş satır kalmadı. " +
                "Excel'de mevcut son satırı aşağı kopyalayarak kapasiteyi artırabilirsiniz.");

        int bosKolon = -1;
        for (int c = h.SipUrunIlkKol; c <= h.SipUrunSonKol; c++)
            if (h.Siparis.Cell(h.SipBaslikSatir, c).GetString().Trim().Length == 0) { bosKolon = c; break; }
        if (bosKolon < 0)
            throw new InvalidOperationException("Sipariş sayfasında yedek ürün sütunu kalmadı.");

        int bosHedKolon = -1;
        for (int c = h.HedUrunIlkKol; c <= h.HedUrunSonKol; c++)
            if (h.Hediye.Cell(h.HedBaslikSatir, c).GetString().Trim().Length == 0) { bosHedKolon = c; break; }
        if (bosHedKolon < 0)
            throw new InvalidOperationException("Hediyelikler sayfasında yedek ürün sütunu kalmadı.");

        h.Stok.Cell(bosSatir, h.StokUrunKol).Value = ad;
        if (baslangicStok > 0) h.Stok.Cell(bosSatir, h.StokBaslangicKol).Value = baslangicStok;
        h.Siparis.Cell(h.SipBaslikSatir, bosKolon).Value = ad;
        h.Hediye.Cell(h.HedBaslikSatir, bosHedKolon).Value = ad;

        Kaydet(wb);
    }

    public void HediyeEkle(string adSoyad, DateTime tarih, IReadOnlyDictionary<string, int> kalemler)
    {
        adSoyad = adSoyad.Trim();
        if (adSoyad.Length == 0) throw new ArgumentException("Ad soyad boş olamaz.");
        if (kalemler.Count == 0 || kalemler.Values.All(v => v <= 0))
            throw new ArgumentException("En az bir ürün için adet girin.");

        using var wb = new XLWorkbook(DosyaYolu);
        var h = SablonHaritasi.Cikar(wb);

        var kolon = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int c = h.HedUrunIlkKol; c <= h.HedUrunSonKol; c++)
        {
            string baslik = h.Hediye.Cell(h.HedBaslikSatir, c).GetString().Trim();
            if (baslik.Length > 0) kolon[baslik] = c;
        }

        foreach (var urun in kalemler.Keys)
            if (!kolon.ContainsKey(urun))
                throw new InvalidOperationException($"\"{urun}\" için Hediyelikler sayfasında sütun bulunamadı.");

        int satir = -1;
        for (int r = h.HedIlkSatir; r <= h.HedSonSatir; r++)
            if (h.Hediye.Cell(r, h.HedKolAd).GetString().Trim().Length == 0) { satir = r; break; }
        if (satir < 0)
            throw new InvalidOperationException(
                "Hediyelikler sayfasında boş satır kalmadı (formüllerin okuduğu aralık doldu).");

        h.Hediye.Cell(satir, h.HedKolAd).Value = adSoyad;
        h.Hediye.Cell(satir, h.HedKolTarih).Value = tarih;
        foreach (var (urun, adet) in kalemler)
            if (adet > 0) h.Hediye.Cell(satir, kolon[urun]).Value = adet;

        Kaydet(wb);
    }

    public void HediyeSil(int satirNo)
    {
        using var wb = new XLWorkbook(DosyaYolu);
        var h = SablonHaritasi.Cikar(wb);

        if (satirNo < h.HedIlkSatir || satirNo > h.HedSonSatir)
            throw new ArgumentOutOfRangeException(nameof(satirNo));

        h.Hediye.Range(satirNo, h.HedKolAd, satirNo, h.HedUrunSonKol)
                .Clear(XLClearOptions.Contents);
        Kaydet(wb);
    }

    public void StokGirisiEkle(DateTime tarih, string urun, int adet, string aciklama)
    {
        if (adet <= 0) throw new ArgumentException("Adet 0'dan büyük olmalı.");

        using var wb = new XLWorkbook(DosyaYolu);
        var h = SablonHaritasi.Cikar(wb);

        int satir = -1;
        for (int r = h.GirIlkSatir; r <= h.GirSonSatir; r++)
            if (h.Giris.Cell(r, h.GirKolUrun).GetString().Trim().Length == 0 &&
                h.Giris.Cell(r, h.GirKolTarih).GetString().Trim().Length == 0)
            { satir = r; break; }
        if (satir < 0)
            throw new InvalidOperationException(
                "Stok Giriş sayfasında boş satır kalmadı (formüllerin okuduğu aralık doldu).");

        h.Giris.Cell(satir, h.GirKolTarih).Value = tarih;
        h.Giris.Cell(satir, h.GirKolUrun).Value = urun;
        h.Giris.Cell(satir, h.GirKolAdet).Value = adet;
        if (!string.IsNullOrWhiteSpace(aciklama)) h.Giris.Cell(satir, h.GirKolAciklama).Value = aciklama;

        Kaydet(wb);
    }

    public void SiparisEkle(string adSoyad, DateTime teslimTarihi, string neden,
                            IReadOnlyDictionary<string, int> kalemler)
    {
        adSoyad = adSoyad.Trim();
        if (adSoyad.Length == 0) throw new ArgumentException("Ad soyad boş olamaz.");
        if (kalemler.Count == 0 || kalemler.Values.All(v => v <= 0))
            throw new ArgumentException("En az bir ürün için adet girin.");

        using var wb = new XLWorkbook(DosyaYolu);
        var h = SablonHaritasi.Cikar(wb);

        var kolon = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int c = h.SipUrunIlkKol; c <= h.SipUrunSonKol; c++)
        {
            string baslik = h.Siparis.Cell(h.SipBaslikSatir, c).GetString().Trim();
            if (baslik.Length > 0) kolon[baslik] = c;
        }

        foreach (var urun in kalemler.Keys)
            if (!kolon.ContainsKey(urun))
                throw new InvalidOperationException($"\"{urun}\" için Sipariş sayfasında sütun bulunamadı.");

        int satir = -1;
        for (int r = h.SipIlkSatir; r <= h.SipSonSatir; r++)
            if (h.Siparis.Cell(r, h.SipKolAd).GetString().Trim().Length == 0) { satir = r; break; }
        if (satir < 0)
            throw new InvalidOperationException(
                "Sipariş sayfasında boş satır kalmadı (formüllerin okuduğu aralık doldu). " +
                "Excel'de mevcut son satırı aşağı kopyalayarak kapasiteyi artırabilirsiniz.");

        h.Siparis.Cell(satir, h.SipKolAd).Value = adSoyad;
        h.Siparis.Cell(satir, h.SipKolTarih).Value = teslimTarihi;
        h.Siparis.Cell(satir, h.SipKolNeden).Value = neden;
        foreach (var (urun, adet) in kalemler)
            if (adet > 0) h.Siparis.Cell(satir, kolon[urun]).Value = adet;

        Kaydet(wb);
    }

    public void SiparisSil(int satirNo)
    {
        using var wb = new XLWorkbook(DosyaYolu);
        var h = SablonHaritasi.Cikar(wb);

        if (satirNo < h.SipIlkSatir || satirNo > h.SipSonSatir)
            throw new ArgumentOutOfRangeException(nameof(satirNo));

        h.Siparis.Range(satirNo, h.SipKolAd, satirNo, h.SipUrunSonKol)
                 .Clear(XLClearOptions.Contents);
        Kaydet(wb);
    }

    public void GirisSil(int satirNo)
    {
        using var wb = new XLWorkbook(DosyaYolu);
        var h = SablonHaritasi.Cikar(wb);

        if (satirNo < h.GirIlkSatir || satirNo > h.GirSonSatir)
            throw new ArgumentOutOfRangeException(nameof(satirNo));

        h.Giris.Range(satirNo, h.GirKolTarih, satirNo, h.GirKolAciklama)
               .Clear(XLClearOptions.Contents);
        Kaydet(wb);
    }

    private static void Kaydet(XLWorkbook wb)
    {

        wb.ForceFullCalculation = true;
        wb.Save();
    }

    private static int HucreSayi(IXLCell hucre)
    {
        if (hucre.TryGetValue<double>(out var d)) return (int)Math.Round(d);
        return 0;
    }

    private static DateTime? HucreTarih(IXLCell hucre)
    {
        if (hucre.TryGetValue<DateTime>(out var t)) return t;
        string s = hucre.GetString().Trim();
        if (s.Length > 0 && DateTime.TryParse(s, new System.Globalization.CultureInfo("tr-TR"),
                System.Globalization.DateTimeStyles.None, out var m))
            return m;
        return null;
    }
}
