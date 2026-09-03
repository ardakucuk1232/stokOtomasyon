using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace StokOtomasyon;

public class StokUrun
{
    public string Ad { get; set; } = "";
    public int Baslangic { get; set; }
    public int Giris { get; set; }
    public int Siparis { get; set; }
    public int Hediye { get; set; }
    public int Kalan => Baslangic + Giris - Siparis - Hediye;

    public const int KritikEsik = 20;

    public string Durum => Kalan <= 0 ? "TÜKENDİ" : (Kalan < KritikEsik ? "KRİTİK" : "YETERLİ");
}

public class SiparisKalemi : INotifyPropertyChanged
{
    public string Urun { get; set; } = "";
    public int Mevcut { get; set; }

    private int _adet;
    public int Adet
    {
        get => _adet;
        set
        {
            if (_adet == value) return;
            _adet = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Adet)));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}

public class KalemBilgi
{
    public string Urun { get; set; } = "";
    public int Adet { get; set; }
}

public class SiparisKaydi
{
    public int SatirNo { get; set; }
    public string Kisi { get; set; } = "";
    public DateTime? Tarih { get; set; }
    public string Neden { get; set; } = "";
    public List<KalemBilgi> Kalemler { get; set; } = new();

    public int ToplamAdet => Kalemler.Sum(k => k.Adet);
    public int CesitSayisi => Kalemler.Count;
    public string TarihYazi => Tarih?.ToString("dd.MM.yyyy") ?? "";
    public string UrunOzeti => string.Join(", ", Kalemler.Select(k => $"{k.Urun} ×{k.Adet}"));
}

public class GirisKaydi
{
    public int SatirNo { get; set; }
    public DateTime? Tarih { get; set; }
    public string Urun { get; set; } = "";
    public int Adet { get; set; }
    public string Aciklama { get; set; } = "";
    public string TarihYazi => Tarih?.ToString("dd.MM.yyyy") ?? "";
}

public class UrunToplami
{
    public string Urun { get; set; } = "";
    public int Toplam { get; set; }
    public double CubukGenisligi { get; set; }
}

public class KisiOzeti
{
    public string Kisi { get; set; } = "";
    public int SiparisSayisi { get; set; }
    public int ToplamAdet { get; set; }
    public DateTime? SonTarih { get; set; }
    public string SonTarihYazi => SonTarih?.ToString("dd.MM.yyyy") ?? "";
}
