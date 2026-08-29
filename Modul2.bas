
Sub Normal_Dagilim_Grafigi()
    Dim z As Double
    Dim sonuc As Double
    Dim i As Integer
    Dim ws As Worksheet
    Dim grafik_sekli As Shape
    Dim grafik_objeleri As Chart

    Set ws = ThisWorkbook.Sheets("Sayfa1")

    ' Veri aralıkları için  sütunları ayarlayalım.
    ws.Range("F1").Value = "Z Değeri"
    ws.Range("G1").Value = "Kümülatif Dağılım(CDF)"

    ' Z değerleri ve kümülatif  dağılım değerlerini hesaplayalım ve bunları yazdıralım
    For i = -30 To 30
        z = i / 10
        sonuc = Application.WorksheetFunction.Norm_S_Dist(z, True)
        ws.Cells(i + 32, 6).Value = z
        ws.Cells(i + 32, 7).Value = sonuc
    Next i

    Set grafik_sekli = ws.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers)
    Set grafik_objeleri = grafik_sekli.Chart

    With grafik_objeleri
        .SetSourceData Source:=ws.Range("F2:G61")
        .ChartType = xlXYScatterSmoothNoMarkers
        .HasTitle = True
        .ChartTitle.Text = "Standart Normal Dağılım"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Z Değeri"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Kümülatif Dağılım"
    End With
End Sub

Sub Standart_Normal_Dagilim_Grafik_Cizimi()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sayfa1") 'Workbook.Sheets'i ws olarak kısalttım ve üstte değişken olarak atadım.

    Dim z As Double
    Dim pdf_degeri As Double
    Dim i As Integer

    ws.Range("D1").Value = "Z Değeri"
    ws.Range("E1").Value = " Olasılık Yoğunluk Fonksiyonu (PDF) Değeri"

    ' Z değerleri ve PDF değerlerini hesapla (-3'ten 3'e, 0.1 aralıklarla oluşturduk.)
    For i = -30 To 30
        z = i / 10
        pdf_degeri = Application.WorksheetFunction.Norm_S_Dist(z, False) ' Olasılık Yoğunluk Fonksiyonu (PDF) için False kullanmalıyız.

        ' Aşağıdaki kodlar bulunan Z ve PDF değerlerini sayfaya yazılacak anlamına geliyor.
        ws.Cells(i + 31, 4).Value = z
        ws.Cells(i + 31, 5).Value = pdf_degeri
    Next i

    Dim grafik_objeleri As ChartObject
    ' Grafik boyutlarını araştırdığım kaynağa göre değiştirdim.Bu sayede görünüm iyileşti.
    Set grafik_objeleri = ws.ChartObjects.Add(Left:=100, Width:=300, Top:=50, Height:=500)

    With grafik_objeleri.Chart
        .SetSourceData Source:=ws.Range("D2:E63")
        .ChartType = xlXYScatterSmoothNoMarkers
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Z Değeri"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Olasılık Yoğunluk Fonksiyonu(PDF)"
        .HasTitle = True
        .ChartTitle.Text = "Standart Normal Dağılım Grafiği"
    End With
End Sub

Sub ListSheetNames()
    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Worksheets
        Debug.Print sht.Name  ' Excel sayfasının ismi vba kodu için önemlidir. Sayfa adında bulamadığım bir  hata aldığım için kontrol amaçlı internetten bulmuş olduğum bu kodu kullandım.
    Next sht
End Sub
