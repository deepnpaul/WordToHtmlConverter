using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
//New Addition On 31-10-2017
using System.IO;
using System.Data;
//=======
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Extensions.Primitives;
using System.Net.Mail;
using Microsoft.AspNetCore.Mvc;
using PuppeteerSharp.Media;
using PuppeteerSharp;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using HtmlToOpenXml;



namespace BKS_REP_ASPCORE.Controllers
{
    public class AktivovPDFController : Controller
    {

        public string BaseSiteURL;
        private readonly IConfiguration _configuration;

        public AktivovPDFController(IConfiguration configuration)
        {
            _configuration = configuration;
            BaseSiteURL = _configuration["BaseSiteURL:DefaultBaseSiteURL"];

        }
        //// GET: AktivovPDF
        //public string BaseSiteURL = System.Configuration.ConfigurationManager.AppSettings["BaseSiteURL"].ToString();
        //public string BaseSiteURL = BaseSiteURL;


        [HttpPost]
        //[ValidateInput(false)]
        [DisableRequestSizeLimit]
        public async Task<FileResult> PortraitAndLandscapePDF_Export(string GridHtmlpdf, string IsLandscape)
        {
            MemoryStream memoryStream = null;
            var browserFetcher = new BrowserFetcher();
            await browserFetcher.DownloadAsync();

            var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            var page = await browser.NewPageAsync();
            if (!string.IsNullOrEmpty(IsLandscape))
            {
                if (IsLandscape == "Landscape")
                {
                    if (!string.IsNullOrEmpty(GridHtmlpdf))
                    {
                        var selfClosingTags = new[] { "img", "br", "input", "hr", "meta", "link", "base", "col", "command", "embed", "keygen", "source", "track", "wbr" };

                        foreach (var tag in selfClosingTags)
                        {
                            GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(GridHtmlpdf, $"<{tag}([^>]*)>", $"<{tag}$1 />", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        }
                        GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(GridHtmlpdf, "<img([^>]*?)src=['\"](.*?)['\"]([^>]*)>", $"<img$1src=\"{BaseSiteURL}$2\"$3>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        string cssStyle = "<style>body { font-size: 12px; } table { width: 100%; table-layout: fixed; } td, th { word-wrap: break-word; }</style>";//+ GridHtmlpdf;

                        int bodyCloseIndex = GridHtmlpdf.IndexOf("</body>", StringComparison.OrdinalIgnoreCase);
                        if (bodyCloseIndex > -1)
                        {
                            GridHtmlpdf = GridHtmlpdf.Insert(bodyCloseIndex, "");
                        }
                        else
                        {
                            GridHtmlpdf = cssStyle + GridHtmlpdf;
                        }
                    }

                    try
                    {
                        await page.SetContentAsync(GridHtmlpdf);

                        var pdfStream = await page.PdfStreamAsync(new PdfOptions
                        {
                            Format = PaperFormat.A4,
                            Landscape = true,
                            MarginOptions = new MarginOptions { Top = "20px", Bottom = "20px", Left = "1cm", Right = "1cm" }
                        });

                        memoryStream = new MemoryStream();
                        await pdfStream.CopyToAsync(memoryStream);
                        memoryStream.Position = 0;

                        //return File(memoryStream.ToArray(), "application/pdf", "Permit-Report-" + DateTime.Now.ToString("MM/dd/yyyy") + ".pdf");
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
                else if (IsLandscape == "Portrait")
                {
                    if (!string.IsNullOrEmpty(GridHtmlpdf))
                    {
                        var selfClosingTags = new[] { "img", "br", "input", "hr", "meta", "link", "base", "col", "command", "embed", "keygen", "source", "track", "wbr" };
                        foreach (var tag in selfClosingTags)
                        {
                            GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(GridHtmlpdf, $"<{tag}([^>]*)>", $"<{tag}$1 />", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        }
                        GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(GridHtmlpdf, "<img([^>]*?)src=['\"](.*?)['\"]([^>]*)>", $"<img$1src=\"{BaseSiteURL}$2\"$3>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(GridHtmlpdf, "<input[^>]*?type=['\"]hidden['\"][^>]*?>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(
                            GridHtmlpdf,
                            "<input(?![^>]*checked=['\"]checked['\"])[^>]*type=['\"]checkbox['\"][^>]*?>",
                            "<div><p>No</p></div>",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase
                        );
                        GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(
                        GridHtmlpdf,
                         "<input[^>]*?checked=['\"]checked['\"][^>]*?type=['\"]checkbox['\"][^>]*?>",
                         "<p>Yes</p>",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(GridHtmlpdf, "(<input[^>]*?type=['\"]radio['\"][^>]*?>)\\s*<label[^>]*?>.*?</label>", "$1", System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);
                        GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(
                            GridHtmlpdf,
                            "<input(?![^>]*checked=['\"]checked['\"])[^>]*type=['\"]radio['\"][^>]*?>",
                            "",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase
                        );
                        GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(
                            GridHtmlpdf,
                            "<input[^>]*?type=['\"]radio['\"][^>]*?checked=['\"]checked['\"][^>]*?>",
                            "<div>Yes</div>",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase
                        );
                        GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(GridHtmlpdf, "<input([^>]*?)value=['\"](.*?)['\"]([^>]*)>", "<div>$2</div>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(GridHtmlpdf, "<textarea[^>]*?>(.*?)</textarea>", "<div>$1</div>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        GridHtmlpdf = System.Text.RegularExpressions.Regex.Replace(GridHtmlpdf, "<select[^>]*?>.*?<option[^>]*?selected=['\"]?selected['\"]?[^>]*?>(.*?)</option>.*?</select>", "<div>$1</div>", System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);

                        string cssStyle = "<style>body { font-size: 12px; margin: 0; padding: 0; } table { width: 100%; table-layout: fixed; } td, th { word-wrap: break-word; } .container { padding-left: 0px; padding-right: 0px; }</style><div class='container'>" + GridHtmlpdf + "</div>";

                        int bodyCloseIndex = GridHtmlpdf.IndexOf("</body>", StringComparison.OrdinalIgnoreCase);
                        if (bodyCloseIndex > -1)
                        {
                            GridHtmlpdf = GridHtmlpdf.Insert(bodyCloseIndex, "");
                        }
                        else
                        {
                            GridHtmlpdf = cssStyle;//+ GridHtmlpdf;
                        }
                    }
                    try
                    {
                        //var headerTemplate = "<div style='font-size:20px; width: 100%; text-align: center; height: 200px;'>Header Content</div><hr>";
                        //var footerTemplate = "<hr><div style='font-size:20px; width: 100%; text-align: center; height: 150px;'>Powered By AKTIVOV</div>";


                        await page.SetContentAsync(GridHtmlpdf);

                        var pdfStream = await page.PdfStreamAsync(new PdfOptions
                        {
                            Format = PaperFormat.A4,
                            Landscape = false
                            //DisplayHeaderFooter = true,
                            //HeaderTemplate = headerTemplate,
                            //FooterTemplate = footerTemplate,
                            //MarginOptions = new MarginOptions { Top = "200px", Bottom = "100px" }
                            // MarginOptions = new MarginOptions { Top = "5px", Bottom = "5px", Left = "0cm", Right = "0cm" },
                            //Scale = 0.9m
                        });

                        memoryStream = new MemoryStream();
                        await pdfStream.CopyToAsync(memoryStream);
                        memoryStream.Position = 0;

                        //return File(memoryStream.ToArray(), "application/pdf", "Invoice-" + DateTime.Now.ToString("MM/dd/yyyy") + ".pdf");
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
            return File(memoryStream.ToArray(), "application/pdf", "File--" + DateTime.Now.ToString("MM/dd/yyyy") + ".pdf"); ;
        }


        //[HttpPost]
        //public async Task<IActionResult> ExportExcelWord(string GridHtmlWord,  string Command)
        //{
        //    GridHtmlWord = Regex.Replace(GridHtmlWord, "<img([^>]*?)src=['\"](.*?)['\"]([^>]*)>", $"<img$1src=\"{BaseSiteURL}$2\"$3>", RegexOptions.IgnoreCase);

        //    if (Command == "ExportWord")
        //    {
        //        var fileName = "WordReport" + DateTime.Now.ToString("MMddyyyyhhmmss") + ".docx";
        //        var contentType = "application/vnd.ms-word";

        //        Response.Clear();
        //        Response.Headers.Clear();
        //        Response.Headers.Add("content-disposition", "attachment;filename=" + fileName);
        //        Response.ContentType = contentType;
        //        await Response.Body.WriteAsync(System.Text.Encoding.UTF8.GetBytes(GridHtmlWord));
        //        await Response.Body.FlushAsync();

        //        return new JsonResult(new { Status = "OK" });
        //    }

        //    return BadRequest(new { Status = "Error", Message = "Invalid Command" });
        //}

        [HttpPost]
        [DisableRequestSizeLimit]
        public async Task<FileResult> ExportToWord(string GridHtmlWord, string IsLandscape)
        {
            MemoryStream memoryStream = null;


            if (!string.IsNullOrEmpty(GridHtmlWord))
            {
                var selfClosingTags = new[] { "img", "br", "input", "hr", "meta", "link", "base", "col", "command", "embed", "keygen", "source", "track", "wbr" };
                foreach (var tag in selfClosingTags)
                {
                    GridHtmlWord = Regex.Replace(GridHtmlWord, $"<{tag}([^>]*)>", $"<{tag}$1 />", RegexOptions.IgnoreCase);
                }
                GridHtmlWord = Regex.Replace(GridHtmlWord, "<img([^>]*?)src=['\"](.*?)['\"]([^>]*)>", $"<img$1src=\"{BaseSiteURL}$2\"$3>", RegexOptions.IgnoreCase);
                GridHtmlWord = Regex.Replace(GridHtmlWord, "<input[^>]*?type=['\"]hidden['\"][^>]*?>", "", RegexOptions.IgnoreCase);
                GridHtmlWord = Regex.Replace(
                    GridHtmlWord,
                    "<input(?![^>]*checked=['\"]checked['\"])[^>]*type=['\"]checkbox['\"][^>]*?>",
                    "<div><p>No</p></div>",
                    RegexOptions.IgnoreCase
                );
                GridHtmlWord = Regex.Replace(
                    GridHtmlWord,
                     "<input[^>]*?checked=['\"]checked['\"][^>]*?type=['\"]checkbox['\"][^>]*?>",
                     "<p>Yes</p>",
                    RegexOptions.IgnoreCase);
                GridHtmlWord = Regex.Replace(GridHtmlWord, "(<input[^>]*?type=['\"]radio['\"][^>]*?>)\\s*<label[^>]*?>.*?</label>", "$1", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                GridHtmlWord = Regex.Replace(
                    GridHtmlWord,
                    "<input(?![^>]*checked=['\"]checked['\"])[^>]*type=['\"]radio['\"][^>]*?>",
                    "",
                    RegexOptions.IgnoreCase
                );
                GridHtmlWord = Regex.Replace(
                    GridHtmlWord,
                    "<input[^>]*?type=['\"]radio['\"][^>]*?checked=['\"]checked['\"][^>]*?>",
                    "<div>Yes</div>",
                    RegexOptions.IgnoreCase
                );
                GridHtmlWord = Regex.Replace(GridHtmlWord, "<input([^>]*?)value=['\"](.*?)['\"]([^>]*)>", "<div>$2</div>", RegexOptions.IgnoreCase);
                GridHtmlWord = Regex.Replace(GridHtmlWord, "<textarea[^>]*?>(.*?)</textarea>", "<div>$1</div>", RegexOptions.IgnoreCase);
                GridHtmlWord = Regex.Replace(GridHtmlWord, "<select[^>]*?>.*?<option[^>]*?selected=['\"]?selected['\"]?[^>]*?>(.*?)</option>.*?</select>", "<div>$1</div>", RegexOptions.IgnoreCase | RegexOptions.Singleline);

                string cssStyle = "<style>body { font-size: 12px; margin: 0; padding: 0; } table { width: 100%; table-layout: fixed; } td, th { word-wrap: break-word; } .container { padding-left: 0px; padding-right: 0px; }</style><div class='container'>" + GridHtmlWord + "</div>";

                int bodyCloseIndex = GridHtmlWord.IndexOf("</body>", StringComparison.OrdinalIgnoreCase);
                if (bodyCloseIndex > -1)
                {
                    GridHtmlWord = GridHtmlWord.Insert(bodyCloseIndex, "");
                }
                else
                {
                    GridHtmlWord = cssStyle;//+ GridHtmlWord;
                }
            }
            try
            {
                memoryStream = new MemoryStream();
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document, true))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    var converter = new HtmlConverter(mainPart);
                    var body = mainPart.Document.Body;

                    // Add the HTML content to the document
                    converter.ParseHtml(GridHtmlWord);

                    mainPart.Document.Save();
                }
                memoryStream.Position = 0;
            }
            finally
            {
                if (memoryStream != null)
                {
                    memoryStream.Dispose();
                }
            }
            return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "File-" + DateTime.Now.ToString("MM_dd_yyyy") + ".docx");
        }
    }
}