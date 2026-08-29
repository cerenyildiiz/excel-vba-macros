# İST347 Vize Ödevi - Hesap Tabloları ve Veri Tabanları

Bu repo, **İST347 - Hesap Tabloları ve Veri Tabanları** dersi kapsamında, **2023-2024 güz döneminde** hazırladığım vize ödevinde kullandığım **VBA (Visual Basic for Applications)** kodlarını içermektedir.

Ödev kapsamında, gruplandırılmış veri serisi üzerinden **aritmetik ortalama, kareli ortalama ve standart sapma** hesaplamaları yapılmıştır. Ayrıca, **standart normal dağılımın kümülatif dağılım fonksiyonu (CDF)** ve **olasılık yoğunluk fonksiyonu (PDF)** için grafikler oluşturulmuştur.

Kodları geliştirmeden önce gerekli hesaplamaları **manuel olarak** gerçekleştirdim. Daha sonra aynı hesaplamaları VBA kullanarak kodladım ve elde edilen sonuçları manuel hesaplamalarla karşılaştırarak kontrol ettim.

## Dosyalar

### `Modul1.bas`

Bu modülde, gruplandırılmış veri serisi üzerinden **aritmetik ortalama, kareli ortalama ve standart sapma** hesaplamaları için kullanılan VBA kodları bulunmaktadır.

- `Ortalamalari_Hesapla(seri As Range)` fonksiyonu, **alt sınır, üst sınır ve frekans** bilgilerinden oluşan üç sütunlu veri aralığını kullanarak aritmetik ve kareli ortalamayı hesaplar.
- Fonksiyon birden fazla sonuç döndürdüğü için sonuçları `Array` kullanarak döndürmektedir. Ödevde `Function` kullanılması istendiğinden bu yöntem tercih edilmiştir.
- `Standart_Sapma_Hesapla` fonksiyonu, gruplandırılmış veri üzerinden **örneklem ve popülasyon standart sapmasını** hesaplar. `orneklem_` parametresi ile hesaplanacak standart sapma türü belirlenir.
- `Merkezi_Egilim_ve_Sapma_Hesabı()` `Sub` prosedürü, ilgili hesaplamaları çalıştırır ve sonuçları **Immediate Window**'a yazdırır.

### Örnek Çıktı

```text
Aritmetik Ortalama: 6,63636363636364
Kareli Ortalama: 49
Örneklem Standart Sapma: 2,33549683248457
Popülasyon Standart Sapma: 2,22680885707562


```

### `Modul2.bas`

Bu modülde, Excel’in `Norm_S_Dist`   fonksiyonundan yararlanılarak standart normal dağılımın kümülatif dağılım fonksiyonu (CDF) ve olasılık yoğunluk fonksiyonu (PDF) için grafikler oluşturulmuştur.

* Normal_Dagilim_Grafigi() prosedürü, -3 ile 3 arasındaki Z değerleri için `Norm_S_Dist` fonksiyonunu kullanarak kümülatif dağılım (CDF) değerlerini hesaplar ve bu değerleri grafik üzerinde gösterir.
* Standart_Normal_Dagilim_Grafik_Cizimi() prosedürü, aynı Z değerleri için `Norm_S_Dist` fonksiyonunu kullanarak olasılık yoğunluk (PDF) değerlerini hesaplar ve standart normal dağılım grafiğini oluşturur.
* ListSheetNames() prosedürü ise çalışma kitabındaki sayfa adlarını listelemek ve kod içerisinde kullanılan sayfa adlarının kontrolünü yapmak amacıyla eklenmiştir.
