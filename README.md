# İST347 Vize Ödevi - Hesap Tabloları ve Veri Tabanları

Bu repo İST347 dersi için hazırladığım vize ödevindeki VBA kodlarını içeriyor. Ödevde gruplandırılmış veri serisi üzerinden aritmetik ortalama, kareli ortalama ve standart sapma hesaplattım, ayrıca standart normal dağılım için de grafik çizdirdim.

Kodları yazmadan önce soruyu önce elle (manuel) çözdüm, sonrasında VBA'ya döktüm. Böylece kodun doğru çalışıp çalışmadığını manuel sonuçlarla karşılaştırarak kontrol edebildim.

## Dosyalar

**Modul1.bas** - aritmetik ortalama, kareli ortalama ve standart sapma hesaplamaları burada bulunmaktadır.

- `Ortalamalari_Hesapla(seri As Range)` fonksiyonu, verdiğim aralıktaki (alt sınır - üst sınır - frekans şeklinde 3 sütunlu) veriden aritmetik ve kareli ortalamayı hesaplayıp Array ile döndürüyor. Soruda function kullanılması istendiği için ByRef yerine dizi kullanmayı tercih ettim.
- `Standart_Sapma_Hesapla` fonksiyonu örneklem ya da popülasyon standart sapmasını hesaplıyor, `orneklem_` parametresi True/False olarak ikisi arasında seçim yapmamı sağlıyor.
- `Merkezi_Egilim_ve_Sapma_Hesabı()` Sub'ı ise hepsini çalıştırıp sonuçları Immediate Window'a yazdırıyor.

Çıktı şu şekilde çıkıyor:
```
Aritmetik Ortalama: 6,63636363636364
Kareli Ortalama: 49
Örneklem Standart Sapma: 2,33549683248457
Popülasyon Standart Sapma: 2,22680885707562
```

**Modul2.bas** - bu modülde de Norm_S_Dist fonksiyonunu kullanarak standart normal dağılımın hem kümülatif dağılım (CDF) hem de olasılık yoğunluk fonksiyonu (PDF) grafiklerini çizdirdim.

- `Normal_Dagilim_Grafigi()` -3 ile 3 arasındaki Z değerleri için CDF hesaplayıp grafiğe döküyor.
- `Standart_Normal_Dagilim_Grafik_Cizimi()` aynı mantıkla PDF grafiğini çiziyor.
- `ListSheetNames()` sayfa adında hata aldığım için kontrol amaçlı ekledim.


## Nasıl çalıştırılır

```
1. Excel > Alt+F11                     # VBA ekranının açılması için bu adım uygulanmaktadır.
2. Insert > Module                     # Yeni modül oluşturulması için kullanılmaktadır.
3. Modul1.bas / Modul2.bas içeriği     # Kodların modül içine aktarılması için yapıştırma
   (ya da File > Import File)            ya da doğrudan içe aktarma işlemi kullanılmaktadır.
4. "Sayfa1"                            # Kod içindeki sayfa adının kendi sayfa adıyla
                                          eşleşmesi gerekmektedir.
5. Range("A2:C6") gibi aralıklar       # Kendi verisine göre güncellenmesi gerekmektedir.
6. Sub çalıştırma: F5                  # Sonuçların görüntülenmesi için Immediate Window
   (Ctrl+G)                              (Ctrl+G) kullanılmaktadır.
```

## Not

Manuel bulduğum sonuçlarla kodun verdiği sonuçlar birebir aynı çıktı, yani kod doğru çalışıyor diyebilirim.

