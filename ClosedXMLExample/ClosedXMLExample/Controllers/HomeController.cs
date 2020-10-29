using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using ClosedXMLExample.Models;
using ClosedXML.Excel;
using System.IO;
using System.Drawing;
using Microsoft.AspNetCore.Http;
using System.Data;
using Newtonsoft.Json;
using ClosedXML.Report;
using Microsoft.AspNetCore.Hosting;
using System;

namespace ClosedXMLExample.Controllers {
    public class HomeController : Controller {
        private readonly ILogger<HomeController> _logger;
        private IHostingEnvironment _env;

        public HomeController(ILogger<HomeController> logger, IHostingEnvironment env) {
            _logger = logger;
            _env = env;
        }

        public IActionResult Index() {
            return View();
        }

        public IActionResult Privacy() {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error() {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        #region ClosedXML Export (excel dışa aktarma) örnekleri
        // excel export örneği, stiller, formatlar ve formüller
        [HttpGet]
        [Route("export/simple")]
        public IActionResult ExportSimpleExcel() {
            byte[] excelFile;

            // Excel dosyası (excel çalışma kitabı) oluşturuyoruz 
            using (var workbook = new XLWorkbook()) {
                // Çalışma kitabına bir sayfa ekliyoruz, sayfa ismini istediğimiz gibi verebiliriz
                var worksheet = workbook.Worksheets.Add("Sayfa 1");

                // A1 hücresine değer atıyoruz 
                worksheet.Cell("A1").Value = "Hello world";

                // Hücrelere alternatif olarak bu şekilde de değer atanabilir, ilk parametre satır index, ikinci parametre sütun index
                worksheet.Cell(1, 3).SetValue(100); // C1
                worksheet.Cell(2, 3).SetValue("200"); // C2

                // Hücrenin veri tipini belirtiyoruz, C1 
                worksheet.Cell(1, 3).DataType = XLDataType.Number;

                // Birden fazla hücreye tek seferde bu şekilde de veri tipi verilebilir, C sütunu ilk 4 satır
                worksheet.Range(worksheet.Cell(1, 3), worksheet.Cell(4, 3)).DataType = XLDataType.Number;

                // Hücreyi biçimlendiriyoruz, binlik ayraçları ve virgülden sonraki hane sayısını belirtmiş olduk, C1
                worksheet.Cell(1, 3).Style.NumberFormat.Format = "#,##0.00";

                // SetFormat metodu ile de aynı işlem yapılabilir, C sütunu ilk 4 satır
                worksheet.Range(worksheet.Cell(1, 3), worksheet.Cell(4, 3)).Style.NumberFormat.SetFormat("#,##0.00");

                // C4 hücresine formül tanımlıyoruz, C1 VE C2 yi topla diyoruz
                worksheet.Cell("C4").SetFormulaA1("=SUM(C1:C2)");

                // Bu metod ise sayfadaki tüm formüllerin tekrar hesaplanmasını sağlıyor
                worksheet.RecalculateAllFormulas();

                // Burada iki hücreyi birleştiriyoruz (A4 ve B4)
                IXLRange totalCell = worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 2)).Merge();
                totalCell.Value = "TOPLAM";

                // Hücreye arka plan rengi veriyoruz, FromColor ya da FromArgb fonksiyonu ile de özel renkler oluşturulabilir
                totalCell.Style.Fill.SetBackgroundColor(XLColor.Black);
                // Alternatif kullanım
                //totalCell.Style.Fill.BackgroundColor = XLColor.FromColor(Color.Black); 

                // Hücrenin fontunu kalın yapıyoruz
                totalCell.Style.Font.SetBold();
                // Alternatif kullanım
                // totalCell.Style.Font.Bold = true; 

                // Hücrenin yazı rengini, tipini ve boyutunu belirtiyoruz
                totalCell.Style.Font.SetFontColor(XLColor.FromColor(Color.White));
                totalCell.Style.Font.SetFontName("Tahoma");
                totalCell.Style.Font.SetFontSize(9);

                // Hücredeki veriyi yatay olarak sağa yasladık
                totalCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

                // Hücredeki veriyi dikey olarak ortaladık
                totalCell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                // Metni Kaydır olarak belirttik 
                totalCell.Style.Alignment.SetWrapText();

                // C4 Hücresine kenarlık oluşturduk ve kenarlık rengi belirttik
                worksheet.Cell("C4").Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
                worksheet.Cell("C4").Style.Border.SetOutsideBorderColor(XLColor.Black);

                // 4. Satır, 3. sütun yani C4 ü yatayda sağa yasladık
                worksheet.Cell(4, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

                // C4 hücresini dikeyde ortaladık
                worksheet.Cell("C4").Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                // C4 hücresini bold yaptık
                worksheet.Cell("C4").Style.Font.SetBold();

                // C4 hücresinin rengini html hex color code ile kırmızı olarak belirttik
                worksheet.Cell("C4").Style.Font.SetFontColor(XLColor.FromHtml("#FF0000"));

                // A1 hücresine noktalı alt kenarlık oluşturduk
                worksheet.Cell("A1").Style.Border.SetBottomBorder(XLBorderStyleValues.Dotted);

                // A sütununun genişliğinin içeriğe göre ayarlanmasını sağladık
                worksheet.Column("A").AdjustToContents();

                // 3. Sütunun yani C sütununun genişliğini belirtiyoruz
                worksheet.Column(3).Width = 20;

                // 4. Satırın yüksekliğini belirtiyoruz 
                worksheet.Row(4).Height = 30;


                using (MemoryStream memoryStream = new MemoryStream()) {
                    // Dosyayı bir path belirterek de kaydedebiliriz, biz burada stream olarak kaydedeceğiz 
                    workbook.SaveAs(memoryStream);

                    // Stream olarak kaydettiğiniz dosyayı byte array olarak alıyoruz, bunu client'a response olarak döneceğiz 
                    excelFile = memoryStream.ToArray();
                }
            }

            // Dosyamızı byte array formatında response olarak dönüyoruz
            return Ok(excelFile);

            // Eğer JavaScript değil de direkt actionresult için belirtilen route'u tetikleyerek dosyayı indirmek istiyorsanız bu alternatifi kullanabilirsiniz.
            //return File(excelFile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

        }

        // basit ürün listesi 
        [HttpGet]
        [Route("export/simple-list")]
        public IActionResult ExportSimpleListExcel() {
            byte[] excelFile;

            // Excel dosyası (excel çalışma kitabı) oluşturuyoruz 
            using (var workbook = new XLWorkbook()) {
                CreateAndFillWorksheet(workbook, "Sayfa 1", products);

                using (MemoryStream memoryStream = new MemoryStream()) {
                    // Dosyayı bir path belirterek de kaydedebiliriz, biz burada stream olarak kaydedeceğiz 
                    workbook.SaveAs(memoryStream);

                    // Stream olarak kaydettiğiniz dosyayı byte array olarak alıyoruz, bunu client'a response olarak döneceğiz 
                    excelFile = memoryStream.ToArray();
                }
            }

            //// Dosyamızı byte array formatında response olarak dönüyoruz
            //return Ok(excelFile);

            //Eğer JavaScript değil de direkt actionresult için belirtilen route'u tetikleyerek dosyayı indirmek istiyorsanız bu alternatifi kullanabilirsiniz.
            return File(excelFile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

        }

        public void CreateAndFillWorksheet(IXLWorkbook workbook, string sheetName, List<Product> products) {
            // Çalışma kitabına bir sayfa ekliyoruz, sayfa ismi parametre olarak geliyor
            var worksheet = workbook.Worksheets.Add(sheetName);

            // ilk satıra tablo başlığımızı oluşturuyoruz
            worksheet.Cell(1, 1).SetValue("IDProduct");
            worksheet.Cell(1, 2).SetValue("Marka");
            worksheet.Cell(1, 3).SetValue("Adı");
            worksheet.Cell(1, 4).SetValue("Fiyat");
            worksheet.Cell(1, 5).SetValue("Stok");
            worksheet.Cell(1, 6).SetValue("Toplam Değeri");

            // ilk satıra stil veriyoruz ( CellsUsed sadece dolu hücrelere işlem yapmamızı sağlar) 
            worksheet.Row(1).CellsUsed().Style.Font.SetBold();
            worksheet.Row(1).CellsUsed().Style.Font.SetFontSize(12);
            worksheet.Row(1).CellsUsed().Style.Fill.SetBackgroundColor(XLColor.LightGray);

            // ikinci satırdan itibaren parametreden gelen ürün listemizi veriyoruz 
            worksheet.Cell(2, 1).InsertData(products);

            // 4. ve 5. sütunları yani fiyat ve stok sütunlarını number olarak formatlıyoruz, (2. satırdan itibaren eklenen ürün sayısı kadar)
            worksheet.Column(4).Cells(2, products.Count + 1).SetDataType(XLDataType.Number).Style.NumberFormat.SetFormat("#,##0.00");
            worksheet.Column(5).Cells(2, products.Count + 1).SetDataType(XLDataType.Number).Style.NumberFormat.SetFormat("#,##0.00");


            // fiyat ve stok sütunlarının harf karşılıklarını alıyoruz
            string priceColumnLetter = worksheet.Column(4).ColumnLetter();
            string stockColumnLetter = worksheet.Column(5).ColumnLetter();

            // exceldeki satırlarımızı 2. satırdan itibaren (yani ürün satırlarımızı) dönüyoruz
            int i, n = worksheet.Rows().Count();
            for (i = 2; i <= n; i++) {
                // aktif satıra toplam değeri hesaplayacak formülü oluşturuyoruz. 
                // örneğin =D2*E2  (i= aktif satır)
                string totalFormula = $"={priceColumnLetter}{i}*{stockColumnLetter}{i}";
                worksheet.Cell(i, 6).SetFormulaA1(totalFormula);
            }


            // toplam değer sütununu formatlıyoruz
            worksheet.Column(6).Cells(2, products.Count + 2).SetDataType(XLDataType.Number).Style.NumberFormat.SetFormat("#,##0.00").Font.SetBold();

            //  toplam ürün değerini hesaplayan formülü de ekliyoruz
            string totalColumnLetter = worksheet.Column(6).ColumnLetter();
            string subTotalFormula = $"=SUM({totalColumnLetter}2:{totalColumnLetter}{n})";

            worksheet.Cell(n + 1, 6).SetFormulaA1(subTotalFormula).Style.Font.SetFontSize(12);

            // tüm formülleri hesaplatıyoruz
            worksheet.RecalculateAllFormulas();

            // sütunların içeriğe göre otomatik genişletilmesini sağlıyoruz
            worksheet.Columns().AdjustToContents();
        }

        // marka bazlı sayfa sayfa ürün listesi oluşturma 
        [HttpGet]
        [Route("export/multipage-list")]
        public IActionResult ExportMultiPageListExcel() {
            byte[] excelFile;

            // ürünlerimizi markaya göre grupluyoruz 
            List<IGrouping<string, Product>> groupedProducts = products.GroupBy(t => t.Brand).ToList();

            // Excel dosyası (excel çalışma kitabı) oluşturuyoruz 
            using (var workbook = new XLWorkbook()) {

                // markalarımızı dönüyoruz
                int i, n = groupedProducts.Count;
                for (i = 0; i < n; i++) {
                    CreateAndFillWorksheet(workbook, groupedProducts[i].Key, groupedProducts[i].ToList());
                }

                using (MemoryStream memoryStream = new MemoryStream()) {
                    // Dosyayı bir path belirterek de kaydedebiliriz, biz burada stream olarak kaydedeceğiz 
                    workbook.SaveAs(memoryStream);

                    // Stream olarak kaydettiğiniz dosyayı byte array olarak alıyoruz, bunu client'a response olarak döneceğiz 
                    excelFile = memoryStream.ToArray();
                }
            }

            //// Dosyamızı byte array formatında response olarak dönüyoruz
            //return Ok(excelFile);

            //Eğer JavaScript değil de direkt actionresult için belirtilen route'u tetikleyerek dosyayı indirmek istiyorsanız bu alternatifi kullanabilirsiniz.
            return File(excelFile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

        }


        // örnek ürün listesi
        readonly List<Product> products = new List<Product> {
            new Product { IDProduct = 1, Brand = "Apple", Name = "IPhone XS 256GB Space Gray", Price = 10000, Stock = 10 },
            new Product { IDProduct = 2, Brand = "Apple", Name = "IPhone XS 256GB Gold", Price = 10500, Stock = 5 },
            new Product { IDProduct = 3, Brand = "Samsung", Name = "Galaxy Note 20 Ultra 256GB", Price = 12999, Stock = 9 },
            new Product { IDProduct = 4, Brand = "Huawei", Name = "P40 Pro 256GB", Price = 11999, Stock = 6 },
            new Product { IDProduct = 5, Brand = "Xiaomi", Name = "Mi 10 256GB 5G", Price = 9999, Stock = 2 }
        };
        #endregion

        #region ClosedXML Import (içe aktar - Excelden okuma) örneği

        //excel upload etme ve excelden veri okuma
        [HttpPost]
        [Route("import/excel")]
        public IActionResult ImportExcel(IFormFile file) {
            System.Data.DataTable dt = new System.Data.DataTable();

            // excel dosyamızı stream'e çeviriyoruz
            using (var ms = new MemoryStream()) {
                file.CopyTo(ms);

                // excel dosyamızı streamden okuyoruz
                using (var workbook = new XLWorkbook(ms)) {
                    var worksheet = workbook.Worksheet(1); // sayfa 1

                    // sayfada kaç sütun kullanılmış onu buluyoruz ve sütunları DataTable'a ekliyoruz, ilk satırda sütun başlıklarımız var
                    int i, n = worksheet.Columns().Count();
                    for (i = 1; i <= n; i++) {
                        dt.Columns.Add(worksheet.Cell(1, i).Value.ToString());
                    }

                    // sayfada kaç satır kullanılmış onu buluyoruz ve DataTable'a satırlarımızı ekliyoruz
                    n = worksheet.Rows().Count();
                    for (i = 2; i <= n; i++) {
                        DataRow dr = dt.NewRow();

                        int j, k = worksheet.Columns().Count();
                        for (j = 1; j <= k; j++) {
                            // i= satır index, j=sütun index, closedXML worksheet için indexler 1'den başlıyor, ama datatable için 0'dan başladığı için j-1 diyoruz
                            dr[j - 1] = worksheet.Cell(i, j).Value;
                        }

                        dt.Rows.Add(dr);
                    }
                }
            }

            // tablomuzu json formatına çeviriyoruz
            string json = JsonConvert.SerializeObject(dt);

            return Ok(json);
        }


        #endregion

        #region ClosedXML.Report kullanımı

        // Sipariş oluştur
        public Order CreateAndGetOrder() {
            Order order = new Order {
                IDOrder = 1,
                CustomerName = "ENES TAŞ",
                CustomerAddress = "X MAH. Y SK. NO:Z KEPEZ ANTALYA",
                CustomerPhone = "0 111 222 33 44",
                OrderDate = DateTime.Now
            };

            order.OrderItems.Add(new OrderItem {
                OrderNo = 1,
                ProductBrand = "APPLE",
                ProductName = "IPhone XS 256GB Space Gray",
                Price = 10000,
                Quantity = 2
            });


            order.OrderItems.Add(new OrderItem {
                OrderNo = 2,
                ProductBrand = "APPLE",
                ProductName = "IPhone XS 256GB Gold",
                Price = 10500,
                Quantity = 1
            });

            order.OrderItems.Add(new OrderItem {
                OrderNo = 3,
                ProductBrand = "Xiaomi",
                ProductName = "Mi 10 256GB 5G",
                Price = 9999,
                Quantity = 2
            });

            order.OrderItems.Add(new OrderItem {
                OrderNo = 4,
                ProductBrand = "Huawei",
                ProductName = "P40 Pro 256GB",
                Price = 11999,
                Quantity = 1
            });

            return order;
        }

        // Sipariş raporu export
        [HttpGet]
        [Route("export/order-report")]
        public IActionResult ExportOrderReport() {
            byte[] excelFile;

            var filePath = Path.Combine(_env.ContentRootPath, "template", "OrderTemplate.xlsx");
            var template = new XLTemplate(filePath); // excel şablonumuzu okuyoruz 
            Order order = CreateAndGetOrder(); // siparişimizi oluşturuyoruz
            template.AddVariable(order); // excel şablonumuza siparişi bind ediyoruz
            template.Generate(); // şablonu generate ediyoruz

            using (MemoryStream ms = new MemoryStream()) {
                // Dosyayı bir path belirterek de kaydedebiliriz, biz burada stream olarak kaydedeceğiz 
                template.SaveAs(ms);

                // Stream olarak kaydettiğiniz dosyayı byte array olarak alıyoruz, bunu client'a response olarak döneceğiz 
                excelFile = ms.ToArray();
            }

            //Eğer JavaScript değil de direkt actionresult için belirtilen route'u tetikleyerek dosyayı indirmek istiyorsanız bu alternatifi kullanabilirsiniz.
            return File(excelFile, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

        }
        #endregion


    }
}
