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
    public int StokFormulIlkKol, StokFormulSonKol;

    public IXLWorksheet Siparis = null!;
    public int SipBaslikSatir, SipKolAd, SipKolTarih, SipKolNeden;
    public int SipUrunIlkKol, SipUrunSonKol, SipToplamKol;
    public int SipIlkSatir, SipSonSatir;

    public IXLWorksheet Hediye = null!;
    public int HedBaslikSatir, HedKolAd, HedKolTarih;
    public int HedUrunIlkKol, HedUrunSonKol, HedToplamKol;
    public int HedIlkSatir, HedSonSatir;

    public IXLWorksheet Giris = null!;
    public int GirBaslikSatir, GirKolTarih, GirKolUrun, GirKolAdet, GirKolAciklama;
    public int GirIlkSatir, GirSonSatir;

    private const int BaslikSatirTavani = 250;
    private const int BaslikKolonTavani = 200;
    private const int VeriTavani = 5000;
    private const int BosSatirToleransi = 20;

    private sealed class Izgara
    {
        public readonly IXLWorksheet Ws;
        public readonly int SatirSayisi, KolonSayisi;
        private readonly string[,] _metin;

        public Izgara(IXLWorksheet ws)
        {
            Ws = ws;
            int sonSatir = ws.LastRowUsed()?.RowNumber() ?? 0;
            int sonKolon = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            SatirSayisi = Math.Min(Math.Max(sonSatir, 1), BaslikSatirTavani);
            KolonSayisi = Math.Min(Math.Max(sonKolon, 1), BaslikKolonTavani);

            _metin = new string[SatirSayisi + 1, KolonSayisi + 1];
            for (int r = 1; r <= SatirSayisi; r++)
                for (int c = 1; c <= KolonSayisi; c++)
                    _metin[r, c] = Duz(Metin(ws.Cell(r, c)));
        }

        public (int Satir, int Kolon)? Bul(Func<string, bool> kosul)
        {
            for (int r = 1; r <= SatirSayisi; r++)
                for (int c = 1; c <= KolonSayisi; c++)
                {
                    string s = _metin[r, c];
                    if (s.Length > 0 && kosul(s)) return (r, c);
                }
            return null;
        }

        public int? SatirdaBul(int satir, Func<string, bool> kosul)
        {
            if (satir < 1 || satir > SatirSayisi) return null;
            for (int c = 1; c <= KolonSayisi; c++)
            {
                string s = _metin[satir, c];
                if (s.Length > 0 && kosul(s)) return c;
            }
            return null;
        }

        public string Oku(int satir, int kolon)
        {
            if (satir < 1 || satir > SatirSayisi || kolon < 1 || kolon > KolonSayisi) return "";
            return _metin[satir, kolon];
        }
    }

    private Izgara _stokIzgara = null!, _sipIzgara = null!, _hedIzgara = null!, _girIzgara = null!;

    public static SablonHaritasi Cikar(XLWorkbook wb)
    {
        var h = new SablonHaritasi();
        var izgaralar = wb.Worksheets.ToDictionary(ws => ws, ws => new Izgara(ws));

        foreach (var ws in wb.Worksheets)
        {
            var iz = izgaralar[ws];

            if (h.Hediye == null && iz.Bul(s => s.Contains("HEDİYE VERİLEN")) != null)
            { h.Hediye = ws; h._hedIzgara = iz; continue; }

            if (h.Siparis == null && iz.Bul(s => s.Contains("VERİLEN KİŞİ")) != null)
            { h.Siparis = ws; h._sipIzgara = iz; continue; }

            if (h.Stok == null &&
                iz.Bul(s => s == "ÜRÜN ADI") != null &&
                iz.Bul(s => s.Contains("BAŞLANGIÇ")) != null)
            { h.Stok = ws; h._stokIzgara = iz; continue; }

            if (h.Giris == null &&
                iz.Bul(s => s == "TARİH") != null &&
                iz.Bul(s => s == "ÜRÜN ADI") != null &&
                iz.Bul(s => s == "ADET") != null)
            { h.Giris = ws; h._girIzgara = iz; continue; }
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
        var urun = _stokIzgara.Bul(s => s == "ÜRÜN ADI")
                   ?? throw new InvalidOperationException("Stok sayfasında 'ÜRÜN ADI' başlığı bulunamadı.");
        StokBaslikSatir = urun.Satir;
        StokUrunKol = urun.Kolon;

        StokBaslangicKol = _stokIzgara.SatirdaBul(StokBaslikSatir, s => s.Contains("BAŞLANGIÇ"))
                           ?? throw new InvalidOperationException("Stok sayfasında 'BAŞLANGIÇ STOK' başlığı bulunamadı.");

        StokIlkSatir = StokBaslikSatir + 1;

        int son = StokIlkSatir - 1;
        for (int r = StokIlkSatir; r <= StokIlkSatir + VeriTavani; r++)
        {
            bool urunVar = Metin(Stok.Cell(r, StokUrunKol)).Trim().Length > 0;
            bool formulVar = SatirdaFormulVar(Stok, r, StokUrunKol + 1, StokUrunKol + 10);
            if (urunVar || formulVar) son = r;
            else if (r > son + BosSatirToleransi) break;
        }
        StokSonSatir = Math.Max(son, StokIlkSatir);

        (StokFormulIlkKol, StokFormulSonKol) = FormulBlogu(
            Stok, OrnekSatirBul(Stok, StokUrunKol, StokIlkSatir, StokSonSatir),
            StokUrunKol + 1, _stokIzgara.KolonSayisi);
    }

    private void SiparisiHaritala()
    {
        var ad = _sipIzgara.Bul(s => s.Contains("VERİLEN KİŞİ"))
                 ?? throw new InvalidOperationException("Sipariş sayfasında 'VERİLEN KİŞİ AD SOYAD' başlığı bulunamadı.");
        SipBaslikSatir = ad.Satir;
        SipKolAd = ad.Kolon;

        SipKolTarih = _sipIzgara.SatirdaBul(SipBaslikSatir, s => s.Contains("TESLİM") || s.Contains("TARİH"))
                      ?? throw new InvalidOperationException("Sipariş sayfasında 'TESLİM TARİHİ' başlığı bulunamadı.");
        SipKolNeden = _sipIzgara.SatirdaBul(SipBaslikSatir, s => s.Contains("NEDEN"))
                      ?? throw new InvalidOperationException("Sipariş sayfasında 'NEDENİ' başlığı bulunamadı.");

        SipToplamKol = ToplamKolonu(_sipIzgara, Siparis, SipBaslikSatir, SipKolNeden + 1);
        (SipUrunIlkKol, SipUrunSonKol) =
            UrunKolonlari(_sipIzgara, SipBaslikSatir, SipKolNeden + 1, SipToplamKol);

        (SipIlkSatir, SipSonSatir) = VeriAlani(
            Siparis, SipBaslikSatir + 1, SipKolAd,
            SipToplamKol > 0 ? SipToplamKol : SipUrunSonKol);
    }

    private void HediyeyiHaritala()
    {
        var ad = _hedIzgara.Bul(s => s.Contains("HEDİYE VERİLEN"))
                 ?? throw new InvalidOperationException("Hediyelikler sayfasında 'HEDİYE VERİLEN KİŞİ' başlığı bulunamadı.");
        HedBaslikSatir = ad.Satir;
        HedKolAd = ad.Kolon;

        HedKolTarih = _hedIzgara.SatirdaBul(HedBaslikSatir, s => s.Contains("TARİH"))
                      ?? throw new InvalidOperationException("Hediyelikler sayfasında 'TARİH' başlığı bulunamadı.");

        HedToplamKol = ToplamKolonu(_hedIzgara, Hediye, HedBaslikSatir, HedKolTarih + 1);
        (HedUrunIlkKol, HedUrunSonKol) =
            UrunKolonlari(_hedIzgara, HedBaslikSatir, HedKolTarih + 1, HedToplamKol);

        (HedIlkSatir, HedSonSatir) = VeriAlani(
            Hediye, HedBaslikSatir + 1, HedKolAd,
            HedToplamKol > 0 ? HedToplamKol : HedUrunSonKol);
    }

    private void GirisiHaritala()
    {
        var tarih = _girIzgara.Bul(s => s == "TARİH")
                    ?? throw new InvalidOperationException("Stok Giriş sayfasında 'TARİH' başlığı bulunamadı.");
        GirBaslikSatir = tarih.Satir;
        GirKolTarih = tarih.Kolon;

        GirKolUrun = _girIzgara.SatirdaBul(GirBaslikSatir, s => s == "ÜRÜN ADI")
                     ?? throw new InvalidOperationException("Stok Giriş sayfasında 'ÜRÜN ADI' başlığı bulunamadı.");
        GirKolAdet = _girIzgara.SatirdaBul(GirBaslikSatir, s => s == "ADET")
                     ?? throw new InvalidOperationException("Stok Giriş sayfasında 'ADET' başlığı bulunamadı.");
        GirKolAciklama = _girIzgara.SatirdaBul(GirBaslikSatir, s => s.Contains("AÇIKLAMA")) ?? (GirKolAdet + 1);

        GirIlkSatir = GirBaslikSatir + 1;
        GirSonSatir = SumifAraligindanSonSatir()
                      ?? SonDoluSatir(Giris, GirKolUrun, GirIlkSatir) + 500;
    }

    private static int ToplamKolonu(Izgara iz, IXLWorksheet ws, int baslikSatir, int ilkKol)
    {
        int? basliktan = iz.SatirdaBul(baslikSatir, s => s == "TOPLAM");
        if (basliktan != null) return basliktan.Value;

        for (int c = ilkKol; c <= iz.KolonSayisi; c++)
            for (int r = baslikSatir + 1; r <= baslikSatir + 5; r++)
                if (ws.Cell(r, c).HasFormula) return c;

        return 0;
    }

    private static (int Ilk, int Son) UrunKolonlari(Izgara iz, int baslikSatir, int ilkKol, int toplamKol)
    {
        int son;
        if (toplamKol > 0)
        {
            son = toplamKol - 1;
        }
        else
        {
            son = ilkKol;
            for (int c = ilkKol; c <= iz.KolonSayisi; c++)
                if (iz.Oku(baslikSatir, c).Length > 0) son = c;
        }
        if (son < ilkKol)
            throw new InvalidOperationException($"'{iz.Ws.Name}' sayfasında ürün sütunu bulunamadı.");
        return (ilkKol, son);
    }

    private static (int Ilk, int Son) VeriAlani(IXLWorksheet ws, int ilkSatir, int adKol, int formulKol)
    {
        int son = ilkSatir - 1;
        for (int r = ilkSatir; r <= ilkSatir + VeriTavani; r++)
        {
            bool adVar = Metin(ws.Cell(r, adKol)).Trim().Length > 0;
            bool formulVar = formulKol > 0 && ws.Cell(r, formulKol).HasFormula;
            if (adVar || formulVar) son = r;
            else if (r > son + BosSatirToleransi) break;
        }
        return (ilkSatir, Math.Max(son, ilkSatir));
    }

    private static int OrnekSatirBul(IXLWorksheet ws, int adKol, int ilkSatir, int sonSatir)
    {
        for (int r = ilkSatir; r <= sonSatir; r++)
            if (Metin(ws.Cell(r, adKol)).Trim().Length > 0) return r;
        return ilkSatir;
    }

    private static (int Ilk, int Son) FormulBlogu(IXLWorksheet ws, int ornekSatir,
                                                  int ilkAramaKol, int kolonSiniri)
    {
        int ilk = 0;
        for (int c = ilkAramaKol; c <= kolonSiniri; c++)
            if (ws.Cell(ornekSatir, c).HasFormula) { ilk = c; break; }
        if (ilk == 0) return (0, 0);

        int son = ilk;
        for (int c = ilk; c <= kolonSiniri; c++)
        {
            if (!ws.Cell(ornekSatir, c).HasFormula) break;
            son = c;
        }
        return (ilk, son);
    }

    private int? SumifAraligindanSonSatir()
    {
        if (StokFormulIlkKol <= 0) return null;

        for (int r = StokIlkSatir; r <= Math.Min(StokSonSatir, StokIlkSatir + 5); r++)
        {
            for (int c = StokFormulIlkKol; c <= Math.Max(StokFormulSonKol, StokFormulIlkKol); c++)
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
