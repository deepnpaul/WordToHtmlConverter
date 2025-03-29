namespace WordToHtml.Models
{
    public class WordToHtmlModel
    {
        public int Id { get; set; }
        public string HtmlContent { get; set; }
        public string Name { get; set; }

        public string FileId { get; set; }
        public string ContentStyle {  get; set; }
    }
}
