using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Mvc;
using OpenXmlPowerTools;
using System.Data.SqlClient;
using System.Xml.Linq;
using System.Diagnostics;
using WordToHtml.Models;
using PuppeteerSharp;
using PuppeteerSharp.Media;
using HtmlAgilityPack;




namespace WordToHtml.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        private readonly string connectionString = "Data Source=LAPTOP-11M4COL7\\SQLEXPRESS;Initial Catalog=WordToHTML;Integrated Security=True;Encrypt=False";

        public IActionResult Index()
        {
            List<WordToHtmlModel> documents = new List<WordToHtmlModel>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "SELECT ID, Name FROM HtmlContentTable";
                SqlCommand command = new SqlCommand(query, connection);

                connection.Open();
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        documents.Add(new WordToHtmlModel
                        {
                            Id = reader.GetInt32(0),
                            Name = reader.GetString(1)
                        });
                    }
                }
            }

            return View(documents);
        }

        public IActionResult Editor(int id)
        {
            WordToHtmlModel model = new WordToHtmlModel();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "SELECT HtmlContent FROM HtmlContentTable WHERE ID = @ID";
                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@ID", id);

                connection.Open();
                model.Id = id;
                model.HtmlContent = (string)command.ExecuteScalar();
            }

            return View(model);
        }

        public IActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        public IActionResult UploadHtml(IFormFile file)
        {

            if (file == null || file.Length == 0)
            {
                TempData["msg"] = "No file selected.";
                return RedirectToAction("Index");
            }
            string fileName = file.FileName;
            var tempFilePath = Path.GetTempFileName();
            using (var stream = new FileStream(tempFilePath, FileMode.Create))
            {
                file.CopyTo(stream);
            }

            string htmlvalue = ConvertDocxToHtml(tempFilePath);

            HtmlDocument document = new HtmlDocument();
            document.LoadHtml(htmlvalue);


            // Remove <html> tag but keep its <style> children
            var htmlNode = document.DocumentNode.SelectSingleNode("//html");
            if (htmlNode != null)
            {
                var parentNode = htmlNode.ParentNode;
                foreach (var childNode in htmlNode.ChildNodes.ToList())
                {
                    parentNode.InsertBefore(childNode, htmlNode);
                }
                htmlNode.Remove();
            }
            // Remove <head> tag
            var headNode = document.DocumentNode.SelectSingleNode("//head");
            if (headNode != null)
            {
                var parentNode = headNode.ParentNode;
                foreach (var childNode in headNode.ChildNodes.ToList())
                {
                    if (childNode.Name != "style")
                    {
                        childNode.Remove();
                    }
                }
                foreach (var childNode in headNode.ChildNodes.ToList())
                {
                    parentNode.InsertBefore(childNode, headNode);
                }
                headNode.Remove();

            }

            // Remove <body> tag but keep its children
            var bodyNode = document.DocumentNode.SelectSingleNode("//body");
            if (bodyNode != null)
            {
                var parentNode = bodyNode.ParentNode;
                foreach (var childNode in bodyNode.ChildNodes.ToList())
                {
                    parentNode.InsertBefore(childNode, bodyNode);
                }
                bodyNode.Remove();
            }

            // Detect and insert hidden line at page breaks
            var pageBreakNodes = document.DocumentNode.SelectNodes("//div[@class='page-break']");
            if (pageBreakNodes != null)
            {
                foreach (var pageBreakNode in pageBreakNodes)
                {
                    var hiddenLine = HtmlNode.CreateNode("<div style='display:none;'>--- Page Break ---</div>");
                    pageBreakNode.ParentNode.InsertBefore(hiddenLine, pageBreakNode);
                }
            }

            string modifiedHtml = document.DocumentNode.OuterHtml;

            htmlvalue = $"<div style=\"padding-left: 43px; padding-right: 10px;\">{modifiedHtml}</div>";


            System.IO.File.Delete(tempFilePath);




            ////////////////////////////////
            int recordId;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "INSERT INTO HtmlContentTable (Name,HtmlContent) VALUES (@Name,@HtmlContent); SELECT SCOPE_IDENTITY();";
                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@Name", fileName);
                command.Parameters.AddWithValue("@HtmlContent", htmlvalue);

                connection.Open();
                recordId = Convert.ToInt32(command.ExecuteScalar());
            }

            // Redirect to the WordEdit action with the record ID as a query parameter
            return RedirectToAction("Editor", new { id = recordId });
        }

        
        public IActionResult SaveEdit(WordToHtmlModel obj)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "UPDATE HtmlContentTable SET HtmlContent = @HtmlContent WHERE Id = @Id";
                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@Id", obj.Id);
                command.Parameters.AddWithValue("@HtmlContent", obj.HtmlContent);

                connection.Open();
                command.ExecuteScalar();
            }

            return RedirectToAction("Index");
        }


        public IActionResult ViewDocument(int id)
        {
            string htmlContent;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "SELECT HtmlContent FROM HtmlContentTable WHERE ID = @ID";
                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@ID", id);

                connection.Open();
                htmlContent = (string)command.ExecuteScalar();
            }

            var model = new WordToHtmlModel
            {
                Id = id,
                HtmlContent = htmlContent
            };

            return View(model);
        }



        public IActionResult Delete(int delete_id)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "delete from HtmlContentTable where Id = @Id";
                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@Id", delete_id);

                connection.Open();
                command.ExecuteScalar();
            }

            return RedirectToAction("Index");
        }

        private string ConvertDocxToHtml(string tempFilePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(tempFilePath, true))
            {
                var settings = new HtmlConverterSettings
                {
                    PageTitle = "Converted HTML"
                };

                // Convert the DOCX file to HTML
                XElement html = HtmlConverter.ConvertToHtml(doc, settings);
                string htmlValue = html.ToString(SaveOptions.DisableFormatting);
                return htmlValue;
                // Save the HTML content to the file
                //File.WriteAllText(htmlFilePath, html.ToString(SaveOptions.DisableFormatting));
            }
        }



        public async Task<FileResult> pdf(WordToHtmlModel obj)
        {
            string GridHtmlpdf = obj.HtmlContent;
            MemoryStream memoryStream = null;
            var browserFetcher = new BrowserFetcher();
            await browserFetcher.DownloadAsync();

            var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            var page = await browser.NewPageAsync();

            try
            {
                await page.SetContentAsync(GridHtmlpdf);

                var pdfStream = await page.PdfStreamAsync(new PdfOptions
                {
                    Format = PaperFormat.A4,
                    Landscape = false,
                    DisplayHeaderFooter = true,
                    //FooterTemplate = footerTemplate,
                    MarginOptions = new MarginOptions { Top = "50px", Bottom = "10px", Left = "1px", Right = "1px" },
                    PrintBackground = true
                });

                memoryStream = new MemoryStream();
                await pdfStream.CopyToAsync(memoryStream);
                //memoryStream.Position = 0;

                return File(memoryStream.ToArray(), "application/pdf", "PDFReport" + DateTime.Now.ToString("MMddyyyyhhmmss") + ".pdf");
            }
            finally
            {
                await page.CloseAsync();
                await browser.CloseAsync();

                if (memoryStream != null)
                {
                    memoryStream.Dispose();
                }
            }
        }


        [HttpPost]
        public async Task<FileResult> TextExportPdf(string GridHtmlpdf)
        {
            MemoryStream memoryStream = null;
            var browserFetcher = new BrowserFetcher();
            await browserFetcher.DownloadAsync();

            var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            var page = await browser.NewPageAsync();

            try
            {
                await page.SetContentAsync(GridHtmlpdf);

                var pdfStream = await page.PdfStreamAsync(new PdfOptions
                {
                    Format = PaperFormat.A4,
                    Landscape = true,
                    DisplayHeaderFooter = true,
                    //FooterTemplate = footerTemplate,
                    MarginOptions = new MarginOptions { Top = "50px", Bottom = "10px", Left = "1px", Right = "1px" },
                    PrintBackground = true
                });

                memoryStream = new MemoryStream();
                await pdfStream.CopyToAsync(memoryStream);
                //memoryStream.Position = 0;

                return File(memoryStream.ToArray(), "application/pdf", "PDFReport" + DateTime.Now.ToString("MMddyyyyhhmmss") + ".pdf");
            }
            finally
            {
                await page.CloseAsync();
                await browser.CloseAsync();

                if (memoryStream != null)
                {
                    memoryStream.Dispose();
                }
            }
        }




    }
}
