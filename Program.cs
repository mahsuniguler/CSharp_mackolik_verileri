//  dotnet add package Selenium.WebDriver
//  dotnet add package HtmlAgilityPack
//  dotnet add package EPPlus//  "EPPlus" Version="7.5.3"
//  "HtmlAgilityPack" Version="1.11.72"
//  "Selenium.Support" Version="4.28.0"
//  "Selenium.WebDriver" Version="4.28.0"
using System;
using System.Threading;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using HtmlAgilityPack;
using OfficeOpenXml;

class Program
{
    private static string CleanText(string input)
    {
        return input.Replace("&nbsp;", " ").Replace("\n", "").Replace("\r", "").Trim();
    }
    public static void mackolik_run(string filePath)
    {
        Console.WriteLine("Veri Çekiliyor.");
        var options = new ChromeOptions();
        var service = ChromeDriverService.CreateDefaultService();
        IWebDriver driver = new ChromeDriver(options);
        string url = "https://arsiv.mackolik.com/Genis-Iddaa-Programi";
        driver.Navigate().GoToUrl(url);
        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
        driver.Manage().Window.Maximize();
        // "Hepsi" seçeneğini seç
        var dropdownElement = wait.Until(d => d.FindElement(By.Id("dayId")));
        SelectElement listbox = new SelectElement(dropdownElement);

        // Checkbox'u işaretle
        var checkbox = wait.Until(d => d.FindElement(By.Id("justNotPlayed")));
        checkbox.Click();

        Thread.Sleep(500);
        listbox.SelectByText("Hepsi");
        Thread.Sleep(2000);
        listbox.SelectByText("Hepsi");
        Thread.Sleep(2000);
        // Tabloyu bekle
        string tableClass = "iddaa-oyna-table";
        wait.Until(d => d.FindElement(By.ClassName(tableClass)).Displayed);

        // HTML içeriğini al
        IWebElement tableHtml = driver.FindElement(By.ClassName(tableClass));
        string tableHtmlText = tableHtml.GetAttribute("outerHTML");

        // HtmlAgilityPack ile HTML'yi parse et
        HtmlDocument doc = new HtmlDocument();
        doc.LoadHtml(tableHtmlText);
        driver.Quit();

        // Değiştirme işlemleri
        foreach (var node in doc.DocumentNode.SelectNodes("//text()"))
        {
            // &nbsp; -> "><"
            if (node.InnerText.Contains("> &nbsp;<"))
            {
                node.InnerHtml = node.InnerHtml.Replace("> &nbsp;<", "><");
            }

            // < -> "><"
            if (node.InnerText.Contains("> <"))
            {
                node.InnerHtml = node.InnerHtml.Replace("> <", "><");
            }

            // <b>&nbsp;</b> -> "----------"
            if (node.InnerHtml.Contains("<b>&nbsp;</b>"))
            {
                node.InnerHtml = node.InnerHtml.Replace("<b>&nbsp;</b>", "");
            }

            // &nbsp; -> ""
            if (node.InnerText.Contains("&nbsp;"))
            {
                node.InnerHtml = node.InnerHtml.Replace("&nbsp;", "");
            }

            // <b> or </b> -> ""
            if (node.InnerHtml.Contains("<b>"))
            {
                node.InnerHtml = node.InnerHtml.Replace("<b>", "").Replace("</b>", "");
            }
            if (node.InnerHtml.Contains("amp;"))
            {
                node.InnerHtml = node.InnerHtml.Replace("amp;", "");
            }
        }
        List<string> imageLinks = new List<string>();
        foreach (var tr in doc.DocumentNode.SelectNodes("//tr"))
        {
            var img = tr.SelectSingleNode("td[5]//img");
            if (img != null)
            {
                string src = img.GetAttributeValue("src", "");
                if (src.EndsWith("1.gif") || src.EndsWith("2.gif") || src.EndsWith("3.gif"))
                {
                    imageLinks.Add(src[^5].ToString());
                }
                else
                {
                    imageLinks.Add(src);
                }
            }
            else
            {
                imageLinks.Add("-");
            }
        }
        // Tabloyu seç
        var rows = doc.DocumentNode.SelectNodes("//tr");

        imageLinks.RemoveRange(0, 3); // İlk üç elemanı sil

        using (var package = new ExcelPackage())
        {
            // Get the first worksheet
            var worksheet = package.Workbook.Worksheets.Add("Tablo");
            // Verileri yaz
            for (int rowIndex = 3; rowIndex < rows.Count; rowIndex++)
            {
                var cells = rows[rowIndex].ChildNodes;
                for (int colIndex = 0; colIndex < cells.Count; colIndex++)
                {
                    string cellText = CleanText(cells[colIndex].InnerText);
                    worksheet.Cells[rowIndex - 2, colIndex + 1].Value = cellText;
                }
            }
            int rowIndex2 = 1; // Veri yazımına 2. satırdan başla

            foreach (var link in imageLinks)
            {
                worksheet.Cells[rowIndex2, 5].Value = link; // Resim bağlantısı

                rowIndex2++;
            }
            // Insert a new empty column at the beginning (first column)
            worksheet.InsertColumn(1, 1); // Inserts at index 1 (the first column)

            // Get the last row in column B
            int lastRow = worksheet.Dimension.End.Row;
            // Process each row in column B
            for (int i = 1; i <= lastRow; i++)
            {
                string cellValue = worksheet.Cells[i, 2].Text;

                // Date processing logic
                string currentDate = "";
                if (cellValue.Contains("."))
                {
                    currentDate = cellValue; // Valid date found
                }
                else if (cellValue.Contains(":"))
                {
                    currentDate = worksheet.Cells[i - 1, 1].Text;
                }
                worksheet.Cells[i, 1].Value = currentDate;
            }
            // Delete rows where column B contains dates
            for (int i = lastRow; i >= 1; i--)
            {
                string m_s = worksheet.Cells[i, 10].Text;
                string i_y = worksheet.Cells[i, 9].Text;
                string cellValue = worksheet.Cells[i, 2].Text;
                if (cellValue.Contains(".") || !m_s.Contains("-") || i_y == "-")
                {

                    worksheet.DeleteRow(i);
                }

            }
            // Excel dosyasını kaydet
            FileInfo excelFile = new FileInfo(filePath);
            package.SaveAs(excelFile);
            Console.WriteLine(filePath + " Adlı Excel Dosyası Kaydedildi");
        }
    }
    static void Main(string[] args)
    {
        string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        string desktopPath = Path.Combine(userProfile, "Desktop", "Mackolik_Verileri");
        if (!Directory.Exists(desktopPath))
        {
            Directory.CreateDirectory(desktopPath);
        }

        string tarih = DateTime.Now.ToString("yyyyMMdd"); // Tarih formatını dilediğiniz gibi değiştirebilirsiniz
        string excel_file_name = $"mackolik_verileri{tarih}.xlsx";
        string filePath = Path.Combine(desktopPath, excel_file_name);

        DateTime currentTime = DateTime.Now;
        // Saat 20:00 ile 23:59 arasında mı kontrol et
        if (currentTime.Hour >= 20 && currentTime.Hour <= 23 && !File.Exists(filePath))
        {
            mackolik_run(filePath);
        }
        else if (File.Exists(filePath))
        {
            Console.WriteLine($"{filePath} Dosyası Mevcuttur");
        }
        else
        {
            int kalan_saat = 20 - currentTime.Hour;
            int kalan_dk = 00 - currentTime.Minute;
            int kalan_sn = 0 - currentTime.Second;

            kalan_sn = kalan_dk * 60 + kalan_saat * 60 * 60 + kalan_sn;
            Console.WriteLine($"{kalan_sn/60} DK Kaldı Lütfen Zamanınında Tekrar Çalıştırınız");
            // Console.WriteLine($"{kalan_dk } Kaldı Lütfen Zamanınında Tekrar Çalıştırınız");
            Console.WriteLine("Saat bu aralıkta değil.");
            Thread.Sleep(kalan_sn*1000);
            mackolik_run(filePath);
        }
    }

}
