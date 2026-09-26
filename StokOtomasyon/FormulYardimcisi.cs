using System;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace StokOtomasyon;

public static class FormulYardimcisi
{
    private static readonly Regex HucreReferansi =
        new(@"(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)", RegexOptions.Compiled);

    public static string SatiraTasi(string formul, int kaynakSatir, int hedefSatir)
    {
        if (string.IsNullOrEmpty(formul)) return formul;
        int fark = hedefSatir - kaynakSatir;
        if (fark == 0) return formul;

        var sonuc = new StringBuilder(formul.Length + 8);
        int i = 0;
        while (i < formul.Length)
        {
            char ch = formul[i];

            if (ch == '"' || ch == '\'')
            {
                int j = i + 1;
                while (j < formul.Length && formul[j] != ch) j++;
                int bitis = Math.Min(j + 1, formul.Length);
                sonuc.Append(formul, i, bitis - i);
                i = bitis;
                continue;
            }

            var eslesme = HucreReferansi.Match(formul, i);
            if (eslesme.Success && eslesme.Index == i)
            {
                bool satirSabit = eslesme.Groups[3].Value == "$";
                if (satirSabit || !int.TryParse(eslesme.Groups[4].Value, out int satir))
                {
                    sonuc.Append(eslesme.Value);
                }
                else
                {
                    int yeni = satir + fark;
                    if (yeni < 1) yeni = 1;
                    sonuc.Append(eslesme.Groups[1].Value)
                         .Append(eslesme.Groups[2].Value)
                         .Append(eslesme.Groups[3].Value)
                         .Append(yeni);
                }
                i += eslesme.Length;
                continue;
            }

            sonuc.Append(ch);
            i++;
        }
        return sonuc.ToString();
    }

    public static int KaynakSatirBul(IXLWorksheet ws, int kolon, int hedefSatir,
                                     int veriIlk, int veriSon)
    {
        for (int r = hedefSatir - 1; r >= veriIlk; r--)
            if (ws.Cell(r, kolon).HasFormula) return r;
        for (int r = hedefSatir + 1; r <= veriSon; r++)
            if (ws.Cell(r, kolon).HasFormula) return r;
        return -1;
    }

    public static int SatiriTamamla(IXLWorksheet ws, int hedefSatir, int ilkKol, int sonKol,
                                    int veriIlk, int veriSon)
    {
        if (ilkKol <= 0 || sonKol < ilkKol) return 0;

        int yazilan = 0;
        for (int c = ilkKol; c <= sonKol; c++)
        {
            var hedef = ws.Cell(hedefSatir, c);
            if (hedef.HasFormula) continue;

            int kaynak = KaynakSatirBul(ws, c, hedefSatir, veriIlk, veriSon);
            if (kaynak < 0) continue;

            string kaynakFormul = ws.Cell(kaynak, c).FormulaA1;
            if (string.IsNullOrWhiteSpace(kaynakFormul)) continue;

            hedef.FormulaA1 = SatiraTasi(kaynakFormul, kaynak, hedefSatir);
            yazilan++;
        }
        return yazilan;
    }
}
