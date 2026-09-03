using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace StokOtomasyon;

public class SablonHaritasi
{

    public IXLWorksheet Stok = null!;
    public int StokBaslikSatir, StokUrunKol, StokBaslangicKol;
    public int StokIlkSatir, StokSonSatir;

    public IXLWorksheet Siparis = null!;
    public int SipBaslikSatir, SipKolAd, SipKolTarih, SipKolNeden;
    public int SipUrunIlkKol, SipUrunSonKol;
    public int SipIlkSatir, SipSonSatir;

    public IXLWorksheet Hediye = null!;
    public int HedBaslikSatir, HedKolAd, HedKolTarih;
    public int HedUrunIlkKol, HedUrunSonKol;
    public int HedIlkSatir, HedSonSatir;

    public IXLWorksheet Giris = null!;
    public int GirBaslikSatir, GirKolTarih, GirKolUrun, GirKolAdet, GirKolAciklama;
    public int GirIlkSatir, GirSonSatir;

    private const int AramaSatir = 40;
    private const int AramaKolon = 60;
    private const int VeriTavani = 5000;

    public static SablonHaritasi Cikar(XLWorkbook wb)
    {
        var h = new SablonHaritasi();

        foreach (var ws in wb.Worksheets)
        {
            if (h.Hediye == null && Bul(ws, s => s.Contains("HEDİYE VERİLEN")) != null)
            { h.Hediye = ws; continue; }

            if (h.Siparis == null && Bul(ws, s => s.Contains("VERİLEN KİŞİ")) != null)
            { h.Siparis = ws; continue; }

            if (h.Stok == null &&
                Bul(ws, s => s == "ÜRÜN ADI") != null &&
                Bul(ws, s => s.Contains("BAŞLANGIÇ")) != null)
            { h.Stok = ws; continue; }

            if (h.Giris == null &&
                Bul(ws, s => s == "TARİH") != null &&
                Bul(ws, s => s == "ÜRÜN ADI") != null &&
                Bul(ws, s => s == "ADET") != null)
            { h.Giris = ws; continue; }
        }

        if (h.Stok == null)
            throw new InvalidOperationException(
                "Stok sayfası bulunamadı: 'ÜRÜN ADI' ve 'BAŞLANGIÇ STOK' başlıklarını içeren bir sayfa gerekli.");
        if (h.Siparis == null)
            throw new InvalidOperationException(
                "Sipariş sayfası bulunamadı: 'VERİLEN KİŞİ AD SOYAD' başlığını içeren bir sayfa gerekli.");
        if (h.Hediye == null)
            throw new InvalidOperationException(
                "Hediyelikler sayfası bulunamadı: 'HEDİYE VERİLEN KİŞİ' başlığını içeren bir sayfa gerekli. " +
                "Lütfen güncel Excel şablonunu kullanın.");
        if (h.Giris == null)
            throw new InvalidOperationException(
                "Stok Giriş sayfası bulunamadı: 'TARİH', 'ÜRÜN ADI' ve 'ADET' başlıklarını içeren bir sayfa gerekli.");

        h.StokuHaritala();
        h.SiparisiHaritala();
        h.HediyeyiHaritala();
        h.GirisiHaritala();
        return h;
    }

    private void StokuHaritala()
    {
        var urun = Bul(Stok, s => s == "ÜRÜN ADI")
                   ?? throw new InvalidOperationException("Stok sayfasında 'ÜRÜN ADI' başlığı bulunamadı.");
        StokBaslikSatir = urun.Satir;
        StokUrunKol = urun.Kolon;

        StokBaslangicKol = SatirdaBul(Stok, StokBaslikSatir, s => s.Contains("BAŞLANGIÇ"))
                           ?? throw new InvalidOperationException("Stok sayfasında 'BAŞLANGIÇ STOK' başlığı bulunamadı.");

        StokIlkSatir = StokBaslikSatir + 1;

        int son = StokIlkSatir - 1;
        for (int r = StokIlkSatir; r <= StokIlkSatir + VeriTavani; r++)
        {
            bool urunVar = Metin(Stok.Cell(r, StokUrunKol)).Trim().Length > 0;
            bool formulVar = SatirdaFormulVar(Stok, r, StokUrunKol + 1, StokUrunKol + 8);
            if (urunVar || formulVar) son = r;
            else if (r > son + 20) break;
        }
        StokSonSatir = Math.Max(son, StokIlkSatir);
    }

    private void SiparisiHaritala()
    {
        var ad = Bul(Siparis, s => s.Contains("VERİLEN KİŞİ"))
                 ?? throw new InvalidOperationException("Sipariş sayfasında 'VERİLEN KİŞİ AD SOYAD' başlığı bulunamadı.");
        SipBaslikSatir = ad.Satir;
        SipKolAd = ad.Kolon;

        SipKolTarih = SatirdaBul(Siparis, SipBaslikSatir, s => s.Contains("TESLİM"))
                      ?? throw new InvalidOperationException("Sipariş sayfasında 'TESLİM TARİHİ' başlığı bulunamadı.");
        SipKolNeden = SatirdaBul(Siparis, SipBaslikSatir, s => s.Contains("NEDEN"))
                      ?? throw new InvalidOperationException("Sipariş sayfasında 'NEDENİ' başlığı bulunamadı.");

        int? toplam = SatirdaBul(Siparis, SipBaslikSatir, s => s == "TOPLAM");
        (SipUrunIlkKol, SipUrunSonKol) =
            UrunKolonlari(Siparis, SipBaslikSatir, SipKolNeden + 1, toplam);

        SipIlkSatir = SipBaslikSatir + 1;
        (SipIlkSatir, SipSonSatir) =
            VeriAlani(Siparis, SipIlkSatir, SipKolAd, toplam ?? SipUrunSonKol);
    }

    private void HediyeyiHaritala()
    {
        var ad = Bul(Hediye, s => s.Contains("HEDİYE VERİLEN"))
                 ?? throw new InvalidOperationException("Hediyelikler sayfasında 'HEDİYE VERİLEN KİŞİ' başlığı bulunamadı.");
        HedBaslikSatir = ad.Satir;
        HedKolAd = ad.Kolon;

        HedKolTarih = SatirdaBul(Hediye, HedBaslikSatir, s => s == "TARİH" || s.Contains("TARİH"))
                      ?? throw new InvalidOperationException("Hediyelikler sayfasında 'TARİH' başlığı bulunamadı.");

        int? toplam = SatirdaBul(Hediye, HedBaslikSatir, s => s == "TOPLAM");
        (HedUrunIlkKol, HedUrunSonKol) =
            UrunKolonlari(Hediye, HedBaslikSatir, HedKolTarih + 1, toplam);

        HedIlkSatir = HedBaslikSatir + 1;
        (HedIlkSatir, HedSonSatir) =
            VeriAlani(Hediye, HedIlkSatir, HedKolAd, toplam ?? HedUrunSonKol);
    }

    private void GirisiHaritala()
    {
        var tarih = Bul(Giris, s => s == "TARİH")
                    ?? throw new InvalidOperationException("Stok Giriş sayfasında 'TARİH' başlığı bulunamadı.");
        GirBaslikSatir = tarih.Satir;
        GirKolTarih = tarih.Kolon;

        GirKolUrun = SatirdaBul(Giris, GirBaslikSatir, s => s == "ÜRÜN ADI")
                     ?? throw new InvalidOperationException("Stok Giriş sayfasında 'ÜRÜN ADI' başlığı bulunamadı.");
        GirKolAdet = SatirdaBul(Giris, GirBaslikSatir, s => s == "ADET")
                     ?? throw new InvalidOperationException("Stok Giriş sayfasında 'ADET' başlığı bulunamadı.");
        GirKolAciklama = SatirdaBul(Giris, GirBaslikSatir, s => s.Contains("AÇIKLAMA")) ?? (GirKolAdet + 1);

        GirIlkSatir = GirBaslikSatir + 1;

        GirSonSatir = SumifAraligindanSonSatir() ?? SonDoluSatir(Giris, GirKolUrun, GirIlkSatir) + 500;
    }

    private static (int Ilk, int Son) UrunKolonlari(IXLWorksheet ws, int baslikSatir,
                                                   int ilkKol, int? toplamKol)
    {
        int son;
        if (toplamKol != null)
        {
            son = toplamKol.Value - 1;
        }
        else
        {
            son = ilkKol;
            for (int c = ilkKol; c <= AramaKolon; c++)
                if (Metin(ws.Cell(baslikSatir, c)).Trim().Length > 0)
                    son = c;
        }
        if (son < ilkKol)
            throw new InvalidOperationException($"'{ws.Name}' sayfasında ürün sütunu bulunamadı.");
        return (ilkKol, son);
    }

    private static (int Ilk, int Son) VeriAlani(IXLWorksheet ws, int ilkSatir,
                                                int adKol, int formulKol)
    {
        int son = ilkSatir - 1;
        for (int r = ilkSatir; r <= ilkSatir + VeriTavani; r++)
        {
            bool adVar = Metin(ws.Cell(r, adKol)).Trim().Length > 0;
            bool formulVar = ws.Cell(r, formulKol).HasFormula;
            if (adVar || formulVar) son = r;
            else if (r > son + 20) break;
        }
        return (ilkSatir, Math.Max(son, ilkSatir));
    }

    private int? SumifAraligindanSonSatir()
    {
        for (int r = StokIlkSatir; r <= Math.Min(StokSonSatir, StokIlkSatir + 5); r++)
        {
            for (int c = StokUrunKol + 1; c <= StokUrunKol + 8; c++)
            {
                string f = Stok.Cell(r, c).FormulaA1 ?? "";
                if (f.Length == 0) continue;

                var m = Regex.Match(f, @"SUMIF\([^!]*!\$?[A-Z]+\$?(\d+):\$?[A-Z]+\$?(\d+)",
                                    RegexOptions.IgnoreCase);
                if (m.Success && int.TryParse(m.Groups[2].Value, out int son)) return son;
            }
        }
        return null;
    }

    public static string Duz(string s)
    {
        s = Regex.Replace(s.Replace("\n", " ").Replace("\r", " "), @"\s+", " ").Trim();
        return s.ToUpper(new CultureInfo("tr-TR"));
    }

    private static string Metin(IXLCell hucre)
    {
        if (hucre.HasFormula) return "";
        try { return hucre.GetString(); }
        catch { return ""; }
    }

    private static (int Satir, int Kolon)? Bul(IXLWorksheet ws, Func<string, bool> kosul)
    {
        for (int r = 1; r <= AramaSatir; r++)
            for (int c = 1; c <= AramaKolon; c++)
            {
                string s = Duz(Metin(ws.Cell(r, c)));
                if (s.Length > 0 && kosul(s)) return (r, c);
            }
        return null;
    }

    private static int? SatirdaBul(IXLWorksheet ws, int satir, Func<string, bool> kosul)
    {
        for (int c = 1; c <= AramaKolon; c++)
        {
            string s = Duz(Metin(ws.Cell(satir, c)));
            if (s.Length > 0 && kosul(s)) return c;
        }
        return null;
    }

    private static bool SatirdaFormulVar(IXLWorksheet ws, int satir, int ilkKol, int sonKol)
    {
        for (int c = ilkKol; c <= sonKol; c++)
            if (ws.Cell(satir, c).HasFormula) return true;
        return false;
    }

    private static int SonDoluSatir(IXLWorksheet ws, int kolon, int ilkSatir)
    {
        int son = ilkSatir;
        for (int r = ilkSatir; r <= ilkSatir + VeriTavani; r++)
            if (Metin(ws.Cell(r, kolon)).Trim().Length > 0) son = r;
            else if (r > son + 50) break;
        return son;
    }
}
