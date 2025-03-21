# Mackolik Verileri Çekme ve Excel'e Aktarma

Bu proje, Mackolik sitesinden iddaa programı verilerini otomatik olarak çekip, verileri Excel dosyasına kaydeden bir C# uygulamasıdır. Kullanıcı belirli bir zaman diliminde (21:00 - 23:59) veri çekme işlemini tetikler ve çekilen veriler belirtilen formatta Excel dosyasına yazılır.

## Özellikler

- **Veri Çekme:** Mackolik sitesinin "Genis-Iddaa-Programi" sayfasından verileri çeker.
- **Zaman Dilimi:** Sadece 21:00 ile 23:59 saatleri arasında veri çekme işlemi yapılır.
- **Excel Çıkışı:** Çekilen veriler, Excel formatında kaydedilir.
- **HTML Parsingi:** HtmlAgilityPack kullanarak HTML verisi işlenir ve temizlenir.
- **Veri Temizleme:** Gereksiz karakterler temizlenir ve veriler doğru formatta yazılır.

## Kullanım

### Gerekli Paketler

Bu projeyi çalıştırabilmek için aşağıdaki NuGet paketlerine ihtiyacınız olacaktır:

- **Selenium.WebDriver** (Versiyon: 4.28.0)
- **Selenium.Support** (Versiyon: 4.28.0)
- **HtmlAgilityPack** (Versiyon: 1.11.72)
- **EPPlus** (Versiyon: 7.5.3)

### Kurulum

1. Projeyi kendi bilgisayarınıza klonlayın.

   ```bash
   git clone https://github.com/mahsuniguler/CSharp_mackolik_verileri.git
   cd CSharp_mackolik_verileri

2. NuGet paketlerini yükleyin:
   ```bash
   dotnet restore

3. Projeyi çalıştırın:
   ```bash
   dotnet run