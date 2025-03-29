using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OpenXmlPowerTools;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using WordToHtml.Models;

namespace WordToHtml.Controllers
{
    public class DocxController : Controller
    {
        private readonly ILogger<DocxController> _logger;

        public DocxController(ILogger<DocxController> logger)
        {
            _logger = logger;
        }

        private readonly string connectionString = "Data Source=LAPTOP-11M4COL7\\SQLEXPRESS;Initial Catalog=WordToHTML;Integrated Security=True;Encrypt=False";

        public IActionResult Index(int id, string Command, WordToHtmlModel obj, IFormFile file)
        {
            WordToHtmlModel doc = new WordToHtmlModel();

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
                ViewBag.documents = documents;
            }

            if (Command == "Edit")
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

            if (Command == "Delete")
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query = "delete from HtmlContentTable where Id = @Id";
                    SqlCommand command = new SqlCommand(query, connection);
                    command.Parameters.AddWithValue("@Id", id);

                    connection.Open();
                    command.ExecuteScalar();
                }
                return RedirectToAction("Index", "Docx");
            }

            if (Command == "EditSave")
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
                return RedirectToAction("Index", "Docx", new { Command = "Edit", id = id });
                //return RedirectToAction("Index", "Docx");
            }

            if (Command == "AddNew")
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
                return RedirectToAction("Index", "Docx", new { Command = "Edit", id = recordId });
                // Redirect to the WordEdit action with the record ID as a query parameter
                //return RedirectToAction("Index","Docx");
            }

            return View(doc);
        }


        public IActionResult View(int id, string Command, WordToHtmlModel obj)
        {
            string Value = "<style>.content {\r\n    max-width: 800px;\r\n    margin: 50px auto;\r\n    padding: 20px;\r\n    background-color: #ffffff;\r\n    box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);\r\n    border-radius: 8px;\r\n}\r\n\r\nh1 {\r\n    text-align: center;\r\n    color: #333333;\r\n}\r\n\r\np {\r\n    margin: 20px 0;\r\n    color: #555555;\r\n    text-align: justify;\r\n}</style><div class=\"content\">\r\n        <h1>India is my Nation</h1>\r\n        <p>\r\n            India, officially known as the Republic of India, is a country in South Asia. It is the seventh-largest country by land area and the second-most populous country in the world. Bounded by the Indian Ocean on the south, the Arabian Sea on the southwest, and the Bay of Bengal on the southeast, it shares land borders with Pakistan to the west; China, Nepal, and Bhutan to the north; and Bangladesh and Myanmar to the east.\r\n        </p><!-- New Page -->\r\n        <p>\r\n            The history of India is marked by diverse cultures, languages, and traditions. It is a land of vibrant colors, rich heritage, and profound spirituality. India is known for its unity in diversity, where people from different backgrounds live together in harmony.\r\n        </p><!-- New Page -->\r\n        <p>\r\n            India is my nation and I am proud of its cultural diversity and rich history. From the snow-capped Himalayas in the north to the tropical beaches in the south, India is a land of incredible beauty and cultural richness. The nation’s deep-rooted traditions and modern advancements make it a unique and fascinating country to live in.\r\n        </p>\r\n    </div>";

            // Regular expression to extract the style content
            var styleRegex = new Regex(@"<style>(.*?)<\/style>", RegexOptions.Singleline);
            var match = styleRegex.Match(Value);

            string styleContent = string.Empty;
            if (match.Success)
            {
                styleContent = match.Groups[1].Value.Trim();
            }


            var pages = Regex.Split(Value, @"<!--\s*New\s*Page\s*-->", RegexOptions.IgnoreCase);

            var pagedContent = pages
                .Select(content => $"<div class=\"page\"><main>{content}</main></div>")
                .Aggregate((current, next) => $"{current}{next}");



            obj.HtmlContent = pagedContent;
            return View(obj);
        }

        public IActionResult Views(int id, string Command, WordToHtmlModel obj)
        {
            string Value = "<style>.content { max-width: 800px; margin: 50px auto; padding: 20px; background-color: #ffffff; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); border-radius: 8px; } h1 { text-align: center; color: #333333; } p { margin: 20px 0; color: #555555; text-align: justify; }</style><div class=\"content\"><h1>India is my Nation</h1><p>India, officially known as the Republic of India, is a country in South Asia. It is the seventh-largest country by land area and the second-most populous country in the world. Bounded by the Indian Ocean on the south, the Arabian Sea on the southwest, and the Bay of Bengal on the southeast, it shares land borders with Pakistan to the west; China, Nepal, and Bhutan to the north; and Bangladesh and Myanmar to the east.</p><!-- New Page --><p>The history of India is marked by diverse cultures, languages, and traditions. It is a land of vibrant colors, rich heritage, and profound spirituality. India is known for its unity in diversity, where people from different backgrounds live together in harmony.</p><!-- New Page --><p>India is my nation and I am proud of its cultural diversity and rich history. From the snow-capped Himalayas in the north to the tropical beaches in the south, India is a land of incredible beauty and cultural richness. The nation’s deep-rooted traditions and modern advancements make it a unique and fascinating country to live in.</p></div>";

            // Extract style content
            var styleRegex = new Regex(@"<style>(.*?)<\/style>", RegexOptions.Singleline);
            var match = styleRegex.Match(Value);

            string styleContent = string.Empty;
            if (match.Success)
            {
                styleContent = match.Groups[1].Value.Trim();
            }

            var pages = Regex.Split(Value, @"<!--\s*New\s*Page\s*-->", RegexOptions.IgnoreCase);

            var pagedContent = pages
                .Select(content => $"<div class=\"page\"><div class=\"content\">{content}</div></div>")
                .Aggregate((current, next) => $"{current}{next}");

            obj.HtmlContent = $"<style>{styleContent}</style>{pagedContent}";
            return View(obj);
        }

        public IActionResult A4Views(int id, string Command, WordToHtmlModel obj)
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



            //var styleRegex = new Regex(@"<style>(.*?)<\/style>", RegexOptions.Singleline);
            //var match = styleRegex.Match(Value);

            //string styleContent = string.Empty;
            //string remainingContent = Value;

            //if (match.Success)
            //{
            //    styleContent = match.Groups[1].Value.Trim();
            //    // Remove the <style> tag 
            //    remainingContent = styleRegex.Replace(Value, string.Empty);
            //}

            //// Split When Found <!-- New Page -->
            //var pagesContent = Regex.Split(remainingContent, @"<!--\s*New\s*Page\s*-->");

            //var wrappedContent = pagesContent
            //    .Select(page => $"<page size=\"A4\" style=\"padding: 10px;\">{page}</page>")
            //    .Aggregate((current, next) => $"{current}{next}");


            //obj.ContentStyle = styleContent;
            //obj.HtmlContent = wrappedContent;


            //obj.HtmlContent = Value;

            return View(model);
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





        public IActionResult ManagePageSample()
        {
            ManageSample doc = new ManageSample();

            List<ManageSample> documents = new List<ManageSample>();


            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "SELECT ID, Name,CompanyName,JoiningDate FROM ManageSample";
                SqlCommand command = new SqlCommand(query, connection);

                connection.Open();
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        documents.Add(new ManageSample
                        {
                            ID = reader.GetGuid(0).ToString(),
                            Name = reader.GetString(1),
                            CompanyName = reader.GetString(2),
                            JoiningDate = reader.IsDBNull(3) ? (DateTime?)null : reader.GetDateTime(3)
                        });
                    }
                }
                ViewBag.documents = documents;
            }
            return View();
        }

    }
}
