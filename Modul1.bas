Function Ortalamalari_Hesapla(seri As Range) As Variant
    Dim toplam_eleman As Double
    Dim toplam_eleman_kareleri As Double
    Dim toplam_frekans As Double
    Dim i As Integer

    toplam_eleman = 0          'İşlemler öncesi 0 ile başlamamız gerektiğinden  toplam değerlerini sıfırladık.
    toplam_eleman_kareleri = 0
    toplam_frekans = 0

    For i = 1 To seri.Rows.Count
        Dim orta_nokta As Double
        orta_nokta = (seri.Cells(i, 1).Value + seri.Cells(i, 2).Value) / 2
        ' Sınıf aralığının orta noktasını bulduk  ve frekanslarla çarptık ve en son bu değerleri topladık.
        toplam_eleman = toplam_eleman + orta_nokta * seri.Cells(i, 3).Value
        toplam_eleman_kareleri = toplam_eleman_kareleri + (orta_nokta ^ 2) * seri.Cells(i, 3).Value
        toplam_frekans = toplam_frekans + seri.Cells(i, 3).Value
    Next i

    If toplam_frekans <> 0 Then 'Eğer toplam frekans 0'dan farklıysa if kısmının içindeki kod çalışır.Bizim durumumuzda 0'dan büyüktür (11>0)
        Ortalamalari_Hesapla = Array(toplam_eleman / toplam_frekans, toplam_eleman_kareleri / toplam_frekans)
    Else
        Ortalamalari_Hesapla = Array(0, 0) ' Dizi kullanarak toplamı frekans toplamına bölerek aritmetik ortalamayı hesapladık. Virgül ile iki farklı ortalamanın (kareli ve aritmetik ortalama) değerlerini elde ettik.
    End If
End Function
Function Standart_Sapma_Hesapla(seri As Range, Optional orneklem_ As Boolean = True) As Double
    Dim toplam_varyans As Double
    Dim toplam_frekans As Double
    Dim ortalamalar As Variant
    Dim aritmetik_ortalama As Double
    Dim i As Integer

    toplam_varyans = 0
    toplam_frekans = 0
    ortalamalar = Ortalamalari_Hesapla(seri)
    aritmetik_ortalama = ortalamalar(0)

    For i = 1 To seri.Rows.Count
        toplam_varyans = toplam_varyans + (seri.Cells(i, 3).Value * (((seri.Cells(i, 1).Value + seri.Cells(i, 2).Value) / 2) - aritmetik_ortalama) ^ 2)
        toplam_frekans = toplam_frekans + seri.Cells(i, 3).Value
    Next i

    If toplam_frekans > 1 Then
        If orneklem_ Then
            Standart_Sapma_Hesapla = Sqr(toplam_varyans / (toplam_frekans - 1))
        Else
            Standart_Sapma_Hesapla = Sqr(toplam_varyans / toplam_frekans)
        End If
    Else
        Standart_Sapma_Hesapla = 0
    End If
End Function

Sub Merkezi_Egilim_ve_Sapma_Hesabı()
    Dim seri_range As Range
    Set seri_range = ThisWorkbook.Sheets("Sayfa1").Range("A2:C6")

    Dim sonuclar As Variant
    sonuclar = Ortalamalari_Hesapla(seri_range)

    Dim standart_sapma As Double
    Dim populasyon_standart_sapma As Double

    ' Örneklem standart sapması hesapla (varsayılan)
    standart_sapma = Standart_Sapma_Hesapla(seri_range)

    ' Popülasyon standart sapması hesaplayabilirsiniz. (isteğe bağlı olarak sonuçlayabilirsiniz.)
    populasyon_standart_sapma = Standart_Sapma_Hesapla(seri_range, False)

    ' Sonuçları Immediate Window'a yazdıralım.
    Debug.Print "Sayfa1"
    Debug.Print "Aritmetik Ortalama: " & sonuclar(0) 'DİPNOT:Soruda koşul olarak aritmetik ortalama ve kareli ortalama bölümünde bir adet function yordam kullanılması istendiğinden bu kodu dizi kullanarak yazdım.ByRef ile de yapılabilirdi ama Sub yerine Function kullanmak sorunun şartıdır bu sebepten dizi konusunda araştırma yapıp onu kullandım.
    Debug.Print "Kareli Ortalama: " & sonuclar(1)
    Debug.Print "Örneklem Standart Sapma: " & standart_sapma
    Debug.Print "Popülasyon Standart Sapma: " & populasyon_standart_sapma
End Sub
