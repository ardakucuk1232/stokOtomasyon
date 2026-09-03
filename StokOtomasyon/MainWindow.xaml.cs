using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;

namespace StokOtomasyon;

public partial class MainWindow : Window
{
    private ExcelService? _excel;
    private bool _hazir;

    private List<StokUrun> _stok = new();
    private List<SiparisKaydi> _siparisler = new();
    private List<SiparisKaydi> _hediyeler = new();
    private List<GirisKaydi> _girisler = new();

    private readonly ObservableCollection<StokUrun> _stokGoster = new();
    private readonly ObservableCollection<SiparisKalemi> _siparisKalemleri = new();
    private readonly ObservableCollection<SiparisKalemi> _siparisGoster = new();
    private readonly ObservableCollection<SiparisKalemi> _hediyeKalemleri = new();
    private readonly ObservableCollection<SiparisKalemi> _hediyeGoster = new();
    private readonly ObservableCollection<SiparisKaydi> _gecmisGoster = new();
    private readonly ObservableCollection<SiparisKaydi> _hediyeGecmisGoster = new();
    private readonly ObservableCollection<GirisKaydi> _girisGoster = new();

    private static string ConfigYolu => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "StokOtomasyon", "config.txt");

    public MainWindow()
    {
        InitializeComponent();
        GridStok.ItemsSource = _stokGoster;
        GridSiparis.ItemsSource = _siparisGoster;
        GridHediye.ItemsSource = _hediyeGoster;
        GridHediyeGecmis.ItemsSource = _hediyeGecmisGoster;
        GridGecmis.ItemsSource = _gecmisGoster;
        GridGiris.ItemsSource = _girisGoster;
        SipTarih.SelectedDate = DateTime.Today;
        HedTarih.SelectedDate = DateTime.Today;
        GirTarih.SelectedDate = DateTime.Today;
        _hazir = true;

        try
        {
            if (File.Exists(ConfigYolu))
            {
                string yol = File.ReadAllText(ConfigYolu).Trim();
                if (File.Exists(yol)) DosyaAc(yol);
            }
        }
        catch {  }

        if (_excel == null)
            TxtDosya.Text = "Başlamak için sağ üstten Excel dosyasını seçin.";
    }

    private void BtnDosyaSec_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "Stok Takip Excel dosyasını seçin",
            Filter = "Excel Dosyası (*.xlsx)|*.xlsx"
        };
        if (dlg.ShowDialog() == true)
            DosyaAc(dlg.FileName);
    }

    private void DosyaAc(string yol)
    {
        try
        {
            _excel = new ExcelService(yol);
            TxtDosya.Text = yol;
            Directory.CreateDirectory(Path.GetDirectoryName(ConfigYolu)!);
            File.WriteAllText(ConfigYolu, yol);
            Yenile();
        }
        catch (Exception ex)
        {
            _excel = null;
            Hata("Dosya açılamadı", ex);
        }
    }

    private void Yenile()
    {
        if (_excel == null) return;
        try
        {
            _stok = _excel.StokOku();
            _siparisler = _excel.SiparisleriOku()
                                .OrderByDescending(s => s.Tarih ?? DateTime.MinValue)
                                .ThenByDescending(s => s.SatirNo)
                                .ToList();
            _hediyeler = _excel.HediyeleriOku()
                               .OrderByDescending(s => s.Tarih ?? DateTime.MinValue)
                               .ThenByDescending(s => s.SatirNo)
                               .ToList();
            _girisler = _excel.GirisleriOku()
                              .OrderByDescending(g => g.Tarih ?? DateTime.MinValue)
                              .ThenByDescending(g => g.SatirNo)
                              .ToList();

            StokListesiniGoster();
            KartlariGuncelle();
            SiparisFormunuHazirla();
            HediyeFormunuHazirla();
            KisiListeleriniGuncelle();
            GecmisiGoster();
            HediyeGecmisiniGoster();
            GirisleriGoster();
            RaporuGuncelle();
        }
        catch (Exception ex)
        {
            Hata("Veriler okunamadı", ex);
        }
    }

    private void BtnYenile_Click(object sender, RoutedEventArgs e) => Yenile();

    private void Sekmeler_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_hazir && e.OriginalSource == Sekmeler && _excel != null) Yenile();
    }

    private void StokListesiniGoster()
    {
        if (!_hazir) return;
        string filtre = (TxtAra.Text ?? "").Trim();
        bool sadeceDusuk = ChkSadeceDusuk.IsChecked == true;

        _stokGoster.Clear();
        foreach (var u in _stok)
        {
            if (filtre.Length > 0 && !u.Ad.Contains(filtre, StringComparison.OrdinalIgnoreCase)) continue;
            if (sadeceDusuk && u.Durum == "YETERLİ") continue;
            _stokGoster.Add(u);
        }
    }

    private void KartlariGuncelle()
    {
        KartUrun.Text = _stok.Count.ToString();
        KartKalan.Text = _stok.Where(u => u.Kalan > 0).Sum(u => u.Kalan).ToString();
        int kritik = _stok.Count(u => u.Durum == "KRİTİK");
        int tukendi = _stok.Count(u => u.Durum == "TÜKENDİ");
        KartKritik.Text = kritik.ToString();
        KartTukendi.Text = tukendi.ToString();

        int dusuk = kritik + tukendi;
        StokSekmeBaslik.Text = dusuk > 0 ? $"📊  Stok Durumu  ⚠{dusuk}" : "📊  Stok Durumu";
    }

    private void TxtAra_TextChanged(object sender, TextChangedEventArgs e) => StokListesiniGoster();
    private void StokFiltreDegisti(object sender, RoutedEventArgs e) => StokListesiniGoster();

    private void SiparisFormunuHazirla()
    {

        var eskiAdet = _siparisKalemleri.ToDictionary(k => k.Urun, k => k.Adet,
                                                      StringComparer.OrdinalIgnoreCase);
        foreach (var k in _siparisKalemleri) k.PropertyChanged -= Kalem_PropertyChanged;
        _siparisKalemleri.Clear();

        foreach (var u in _stok)
        {
            var kalem = new SiparisKalemi
            {
                Urun = u.Ad,
                Mevcut = u.Kalan,
                Adet = eskiAdet.TryGetValue(u.Ad, out int a) ? a : 0
            };
            kalem.PropertyChanged += Kalem_PropertyChanged;
            _siparisKalemleri.Add(kalem);
        }
        SiparisUrunFiltrele();
        CanliOzetiGuncelle();

        string secili = GirUrun.Text;
        GirUrun.ItemsSource = _stok.Select(u => u.Ad).ToList();
        GirUrun.Text = secili;
    }

    private void Kalem_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        => CanliOzetiGuncelle();

    private void CanliOzetiGuncelle()
    {
        var dolu = _siparisKalemleri.Where(k => k.Adet > 0).ToList();
        SipCanliOzet.Text = dolu.Count == 0
            ? "Sepet boş — adet girdikçe özet burada görünür."
            : $"Sepet: {dolu.Count} çeşit, toplam {dolu.Sum(k => k.Adet)} adet   →   " +
              string.Join(", ", dolu.Take(4).Select(k => $"{k.Urun} ×{k.Adet}")) +
              (dolu.Count > 4 ? " …" : "");
    }

    private void SipUrunAra_TextChanged(object sender, TextChangedEventArgs e) => SiparisUrunFiltrele();

    private void SiparisUrunFiltrele()
    {
        if (!_hazir) return;
        string filtre = (SipUrunAra.Text ?? "").Trim();
        _siparisGoster.Clear();
        foreach (var k in _siparisKalemleri)
            if (filtre.Length == 0 || k.Urun.Contains(filtre, StringComparison.OrdinalIgnoreCase)
                || k.Adet > 0)
                _siparisGoster.Add(k);
    }

    private void BtnSiparisTemizle_Click(object sender, RoutedEventArgs e)
    {
        foreach (var k in _siparisKalemleri) k.Adet = 0;
        SipAdSoyad.Text = "";
        SipUrunAra.Text = "";
        SiparisUrunFiltrele();
    }

    private void BtnSiparisKaydet_Click(object sender, RoutedEventArgs e)
    {
        if (!DosyaHazirMi()) return;

        string ad = (SipAdSoyad.Text ?? "").Trim();
        if (ad.Length == 0) { Uyari("Lütfen ad soyad girin."); return; }
        if (SipTarih.SelectedDate == null) { Uyari("Lütfen teslim tarihi seçin."); return; }

        var kalemler = _siparisKalemleri.Where(k => k.Adet > 0)
                                        .ToDictionary(k => k.Urun, k => k.Adet);
        if (kalemler.Count == 0) { Uyari("En az bir ürün için adet girin."); return; }

        var yetersiz = _siparisKalemleri.Where(k => k.Adet > 0 && k.Adet > k.Mevcut)
                                        .Select(k => $"• {k.Urun} (mevcut {k.Mevcut}, istenen {k.Adet})")
                                        .ToList();
        if (yetersiz.Count > 0)
        {
            var cevap = MessageBox.Show(
                "Aşağıdaki ürünlerde stok yetersiz:\n\n" + string.Join("\n", yetersiz) +
                "\n\nYine de kaydedilsin mi? (Stok eksiye düşer)",
                "Stok Yetersiz", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (cevap != MessageBoxResult.Yes) return;
        }

        try
        {
            string neden = (SipNeden.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "SİPARİŞ";
            _excel!.SiparisEkle(ad, SipTarih.SelectedDate.Value, neden, kalemler);
            Bilgi($"Sipariş kaydedildi: {ad} — {kalemler.Sum(k => k.Value)} adet, {kalemler.Count} ürün.");
            foreach (var k in _siparisKalemleri) k.Adet = 0;
            SipAdSoyad.Text = "";
            Yenile();
        }
        catch (Exception ex)
        {
            Hata("Sipariş kaydedilemedi", ex);
        }
    }

    private void HediyeFormunuHazirla()
    {
        var eskiAdet = _hediyeKalemleri.ToDictionary(k => k.Urun, k => k.Adet,
                                                     StringComparer.OrdinalIgnoreCase);
        foreach (var k in _hediyeKalemleri) k.PropertyChanged -= HediyeKalem_PropertyChanged;
        _hediyeKalemleri.Clear();

        foreach (var u in _stok)
        {
            var kalem = new SiparisKalemi
            {
                Urun = u.Ad,
                Mevcut = u.Kalan,
                Adet = eskiAdet.TryGetValue(u.Ad, out int a) ? a : 0
            };
            kalem.PropertyChanged += HediyeKalem_PropertyChanged;
            _hediyeKalemleri.Add(kalem);
        }
        HediyeUrunFiltrele();
        HediyeOzetiGuncelle();
    }

    private void HediyeKalem_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        => HediyeOzetiGuncelle();

    private void HediyeOzetiGuncelle()
    {
        var dolu = _hediyeKalemleri.Where(k => k.Adet > 0).ToList();
        HedCanliOzet.Text = dolu.Count == 0
            ? "Adet girdikçe özet burada görünür. Hediyeler stoktan düşer."
            : $"Hediye: {dolu.Count} çeşit, toplam {dolu.Sum(k => k.Adet)} adet   →   " +
              string.Join(", ", dolu.Take(4).Select(k => $"{k.Urun} ×{k.Adet}")) +
              (dolu.Count > 4 ? " …" : "");
    }

    private void HedUrunAra_TextChanged(object sender, TextChangedEventArgs e) => HediyeUrunFiltrele();

    private void HediyeUrunFiltrele()
    {
        if (!_hazir) return;
        string filtre = (HedUrunAra.Text ?? "").Trim();
        _hediyeGoster.Clear();
        foreach (var k in _hediyeKalemleri)
            if (filtre.Length == 0 || k.Urun.Contains(filtre, StringComparison.OrdinalIgnoreCase)
                || k.Adet > 0)
                _hediyeGoster.Add(k);
    }

    private void BtnHediyeTemizle_Click(object sender, RoutedEventArgs e)
    {
        foreach (var k in _hediyeKalemleri) k.Adet = 0;
        HedAdSoyad.Text = "";
        HedUrunAra.Text = "";
        HediyeUrunFiltrele();
    }

    private void BtnHediyeKaydet_Click(object sender, RoutedEventArgs e)
    {
        if (!DosyaHazirMi()) return;

        string ad = (HedAdSoyad.Text ?? "").Trim();
        if (ad.Length == 0) { Uyari("Lütfen ad soyad girin."); return; }
        if (HedTarih.SelectedDate == null) { Uyari("Lütfen tarih seçin."); return; }

        var kalemler = _hediyeKalemleri.Where(k => k.Adet > 0)
                                       .ToDictionary(k => k.Urun, k => k.Adet);
        if (kalemler.Count == 0) { Uyari("En az bir ürün için adet girin."); return; }

        var yetersiz = _hediyeKalemleri.Where(k => k.Adet > 0 && k.Adet > k.Mevcut)
                                       .Select(k => $"• {k.Urun} (mevcut {k.Mevcut}, istenen {k.Adet})")
                                       .ToList();
        if (yetersiz.Count > 0)
        {
            var cevap = MessageBox.Show(
                "Aşağıdaki ürünlerde stok yetersiz:\n\n" + string.Join("\n", yetersiz) +
                "\n\nYine de kaydedilsin mi? (Stok eksiye düşer)",
                "Stok Yetersiz", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (cevap != MessageBoxResult.Yes) return;
        }

        try
        {
            _excel!.HediyeEkle(ad, HedTarih.SelectedDate.Value, kalemler);
            Bilgi($"Hediye kaydedildi: {ad} — {kalemler.Sum(k => k.Value)} adet, {kalemler.Count} ürün.");
            foreach (var k in _hediyeKalemleri) k.Adet = 0;
            HedAdSoyad.Text = "";
            Yenile();
        }
        catch (Exception ex)
        {
            Hata("Hediye kaydedilemedi", ex);
        }
    }

    private void HediyeGecmisiniGoster()
    {
        _hediyeGecmisGoster.Clear();
        foreach (var s in _hediyeler) _hediyeGecmisGoster.Add(s);
    }

    private void BtnHediyeSil_Click(object sender, RoutedEventArgs e)
    {
        if (!DosyaHazirMi()) return;
        if (GridHediyeGecmis.SelectedItem is not SiparisKaydi s)
        { Uyari("Silmek için listeden bir hediye kaydı seçin."); return; }

        var cevap = MessageBox.Show(
            $"{s.Kisi} — {s.TarihYazi} tarihli hediye kaydı silinsin mi?\n" +
            $"({s.UrunOzeti})\n\nSilinen adetler stoğa geri eklenir.",
            "Hediyeyi Sil", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (cevap != MessageBoxResult.Yes) return;

        try
        {
            _excel!.HediyeSil(s.SatirNo);
            Bilgi("Hediye kaydı silindi, adetler stoğa geri döndü.");
            Yenile();
        }
        catch (Exception ex)
        {
            Hata("Hediye silinemedi", ex);
        }
    }

    private void KisiListeleriniGuncelle()
    {
        var kisiler = _siparisler.Select(s => s.Kisi)
                                 .Concat(_hediyeler.Select(s => s.Kisi))
                                 .Distinct(StringComparer.OrdinalIgnoreCase)
                                 .OrderBy(k => k)
                                 .ToList();

        string sipYazi = SipAdSoyad.Text, gecYazi = GecKisi.Text, hedYazi = HedAdSoyad.Text;
        SipAdSoyad.ItemsSource = kisiler;
        GecKisi.ItemsSource = kisiler;
        HedAdSoyad.ItemsSource = kisiler;
        SipAdSoyad.Text = sipYazi;
        GecKisi.Text = gecYazi;
        HedAdSoyad.Text = hedYazi;
    }

    private void GecmisiGoster()
    {
        if (!_hazir) return;

        string kisi = (GecKisi.Text ?? "").Trim();
        DateTime? bas = GecBas.SelectedDate, bit = GecBit.SelectedDate;
        string neden = (GecNeden.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "TÜMÜ";

        _gecmisGoster.Clear();
        foreach (var s in _siparisler)
        {
            if (kisi.Length > 0 && !s.Kisi.Contains(kisi, StringComparison.OrdinalIgnoreCase)) continue;
            if (bas != null && (s.Tarih == null || s.Tarih < bas)) continue;
            if (bit != null && (s.Tarih == null || s.Tarih > bit)) continue;
            if (neden != "TÜMÜ" && !string.Equals(s.Neden, neden, StringComparison.OrdinalIgnoreCase)) continue;
            _gecmisGoster.Add(s);
        }

        GecSayac.Text = $"{_gecmisGoster.Count} kayıt listeleniyor (toplam {_siparisler.Count}).";

        if (_gecmisGoster.Count == 0)
        {
            GecOzet.Text = "";
        }
        else
        {
            int toplam = _gecmisGoster.Sum(s => s.ToplamAdet);
            string kim = kisi.Length > 0 ? $"\"{kisi}\" için " : "";
            GecOzet.Text = $"{kim}{_gecmisGoster.Count} sipariş, toplam {toplam} adet.";
        }
    }

    private void GecmisFiltreDegisti(object sender, SelectionChangedEventArgs e) => GecmisiGoster();
    private void GecmisFiltreYaziDegisti(object sender, KeyEventArgs e) => GecmisiGoster();

    private void BtnGecmisSifirla_Click(object sender, RoutedEventArgs e)
    {
        GecKisi.Text = "";
        GecBas.SelectedDate = null;
        GecBit.SelectedDate = null;
        GecNeden.SelectedIndex = 0;
        GecmisiGoster();
    }

    private void GridGecmis_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (GridGecmis.SelectedItem is SiparisKaydi s)
        {
            DetayBaslik.Text = $"{s.Kisi} — {s.TarihYazi} ({s.Neden})";
            GridDetay.ItemsSource = s.Kalemler;
        }
        else
        {
            DetayBaslik.Text = "Sipariş Detayı";
            GridDetay.ItemsSource = null;
        }
    }

    private void BtnSiparisSil_Click(object sender, RoutedEventArgs e)
    {
        if (!DosyaHazirMi()) return;
        if (GridGecmis.SelectedItem is not SiparisKaydi s)
        { Uyari("Silmek için listeden bir sipariş seçin."); return; }

        var cevap = MessageBox.Show(
            $"{s.Kisi} — {s.TarihYazi} tarihli sipariş silinsin mi?\n" +
            $"({s.UrunOzeti})\n\nSilinen adetler stoğa geri eklenir.",
            "Siparişi Sil", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (cevap != MessageBoxResult.Yes) return;

        try
        {
            _excel!.SiparisSil(s.SatirNo);
            Bilgi("Sipariş silindi, adetler stoğa geri döndü.");
            Yenile();
        }
        catch (Exception ex)
        {
            Hata("Sipariş silinemedi", ex);
        }
    }

    private void BtnCsvAktar_Click(object sender, RoutedEventArgs e)
    {
        if (_gecmisGoster.Count == 0) { Uyari("Dışa aktarılacak kayıt yok."); return; }

        var dlg = new SaveFileDialog
        {
            Title = "CSV olarak kaydet",
            Filter = "CSV Dosyası (*.csv)|*.csv",
            FileName = $"siparis_gecmisi_{DateTime.Today:yyyy_MM_dd}.csv"
        };
        if (dlg.ShowDialog() != true) return;

        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("Kişi;Tarih;Nedeni;Ürün;Adet");
            foreach (var s in _gecmisGoster)
                foreach (var k in s.Kalemler)
                    sb.AppendLine($"{Csv(s.Kisi)};{s.TarihYazi};{Csv(s.Neden)};{Csv(k.Urun)};{k.Adet}");

            File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(true));
            Bilgi($"CSV kaydedildi:\n{dlg.FileName}");
        }
        catch (Exception ex)
        {
            Hata("CSV kaydedilemedi", ex);
        }
    }

    private static string Csv(string s) =>
        s.Contains(';') || s.Contains('"')
            ? "\"" + s.Replace("\"", "\"\"") + "\""
            : s;

    private void GirisleriGoster()
    {
        _girisGoster.Clear();
        foreach (var g in _girisler) _girisGoster.Add(g);
    }

    private void BtnGirisKaydet_Click(object sender, RoutedEventArgs e)
    {
        if (!DosyaHazirMi()) return;

        string urun = (GirUrun.Text ?? "").Trim();
        if (urun.Length == 0) { Uyari("Lütfen ürün seçin."); return; }
        if (!_stok.Any(u => string.Equals(u.Ad, urun, StringComparison.OrdinalIgnoreCase)))
        { Uyari($"\"{urun}\" stok listesinde yok. Önce Ürün Ekle sekmesinden ekleyin."); return; }
        if (GirTarih.SelectedDate == null) { Uyari("Lütfen tarih seçin."); return; }
        if (!int.TryParse(GirAdet.Text.Trim(), out int adet) || adet <= 0)
        { Uyari("Adet pozitif bir sayı olmalı."); return; }

        try
        {
            _excel!.StokGirisiEkle(GirTarih.SelectedDate.Value, urun, adet, GirAciklama.Text);
            Bilgi($"Stok girişi kaydedildi: {urun} +{adet}");
            GirAdet.Clear(); GirAciklama.Clear();
            Yenile();
        }
        catch (Exception ex)
        {
            Hata("Stok girişi kaydedilemedi", ex);
        }
    }

    private void BtnGirisSil_Click(object sender, RoutedEventArgs e)
    {
        if (!DosyaHazirMi()) return;
        if (GridGiris.SelectedItem is not GirisKaydi g)
        { Uyari("Silmek için listeden bir giriş seçin."); return; }

        var cevap = MessageBox.Show(
            $"{g.TarihYazi} — {g.Urun} +{g.Adet} girişi silinsin mi?\nStok bu adet kadar azalır.",
            "Girişi Sil", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (cevap != MessageBoxResult.Yes) return;

        try
        {
            _excel!.GirisSil(g.SatirNo);
            Bilgi("Stok girişi silindi.");
            Yenile();
        }
        catch (Exception ex)
        {
            Hata("Giriş silinemedi", ex);
        }
    }

    private void RaporuGuncelle()
    {
        if (!_hazir) return;

        DateTime? bas = RapBas.SelectedDate, bit = RapBit.SelectedDate;
        var secim = _siparisler.Where(s =>
                (bas == null || (s.Tarih != null && s.Tarih >= bas)) &&
                (bit == null || (s.Tarih != null && s.Tarih <= bit)))
            .ToList();

        RapSiparis.Text = secim.Count.ToString();
        RapAdet.Text = secim.Sum(s => s.ToplamAdet).ToString();

        var urunToplam = secim.SelectMany(s => s.Kalemler)
                              .GroupBy(k => k.Urun, StringComparer.OrdinalIgnoreCase)
                              .Select(g => new { Urun = g.Key, Toplam = g.Sum(k => k.Adet) })
                              .OrderByDescending(x => x.Toplam)
                              .ToList();

        RapUrun.Text = urunToplam.Count > 0 ? $"{urunToplam[0].Urun} ({urunToplam[0].Toplam})" : "—";

        var kisiler = secim.GroupBy(s => s.Kisi, StringComparer.OrdinalIgnoreCase)
                           .Select(g => new KisiOzeti
                           {
                               Kisi = g.Key,
                               SiparisSayisi = g.Count(),
                               ToplamAdet = g.Sum(s => s.ToplamAdet),
                               SonTarih = g.Max(s => s.Tarih)
                           })
                           .OrderByDescending(k => k.ToplamAdet)
                           .ToList();

        RapKisi.Text = kisiler.Count > 0 ? $"{kisiler[0].Kisi} ({kisiler[0].ToplamAdet})" : "—";
        GridKisiler.ItemsSource = kisiler;

        int max = urunToplam.Count > 0 ? urunToplam[0].Toplam : 1;
        if (max <= 0) max = 1;
        ListeTopUrun.ItemsSource = urunToplam.Take(10)
            .Select(x => new UrunToplami
            {
                Urun = x.Urun,
                Toplam = x.Toplam,
                CubukGenisligi = Math.Max(4.0, 420.0 * x.Toplam / max)
            })
            .ToList();
    }

    private void RaporFiltreDegisti(object sender, SelectionChangedEventArgs e) => RaporuGuncelle();

    private void BtnRaporSifirla_Click(object sender, RoutedEventArgs e)
    {
        RapBas.SelectedDate = null;
        RapBit.SelectedDate = null;
        RaporuGuncelle();
    }

    private void BtnUrunEkle_Click(object sender, RoutedEventArgs e)
    {
        if (!DosyaHazirMi()) return;

        string ad = UrnAd.Text.Trim();
        if (ad.Length == 0) { Uyari("Lütfen ürün adı girin."); return; }

        int baslangic = 0;
        if (UrnBaslangic.Text.Trim().Length > 0 &&
            (!int.TryParse(UrnBaslangic.Text.Trim(), out baslangic) || baslangic < 0))
        { Uyari("Başlangıç stok sayı olmalı."); return; }

        try
        {
            _excel!.UrunEkle(ad, baslangic);
            Bilgi($"Ürün eklendi: {ad}");
            UrnAd.Clear(); UrnBaslangic.Clear();
            Yenile();
        }
        catch (Exception ex)
        {
            Hata("Ürün eklenemedi", ex);
        }
    }

    private bool DosyaHazirMi()
    {
        if (_excel == null)
        {
            Uyari("Önce sağ üstten Excel dosyasını seçin.");
            return false;
        }
        return true;
    }

    private static void Uyari(string mesaj) =>
        MessageBox.Show(mesaj, "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);

    private static void Bilgi(string mesaj) =>
        MessageBox.Show(mesaj, "Bilgi", MessageBoxButton.OK, MessageBoxImage.Information);

    private static void Hata(string baslik, Exception ex)
    {
        string ek = ex is IOException
            ? "\n\nİpucu: Dosya Excel'de açıksa kapatıp tekrar deneyin."
            : "";
        MessageBox.Show($"{baslik}:\n{ex.Message}{ek}", "Hata",
                        MessageBoxButton.OK, MessageBoxImage.Error);
    }
}
