using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Utilities;
using System.Net;
using System.Text.RegularExpressions;

namespace DocumentProcessor.Services
{
    public class HtmlToWordConverter : IHtmlToWordConverter
    {
        public string ConvertHtmlToWordFormat(string html)
        {
            if (string.IsNullOrEmpty(html))
                return string.Empty;

            // Remove HTML tags and convert common elements
            html = Regex.Replace(html, @"<br\s*/>", "\n");
            html = Regex.Replace(html, @"<p.*?>", "");
            html = Regex.Replace(html, @"</p>", "\n");
            html = Regex.Replace(html, @"<div.*?>", "");
            html = Regex.Replace(html, @"</div>", "\n");
            html = Regex.Replace(html, @"<span.*?>", "");
            html = Regex.Replace(html, @"</span>", "");

            // Convert HTML entities
            html = WebUtility.HtmlDecode(html);

            // Remove any remaining HTML tags
            html = Regex.Replace(html, @"<[^>]+>", string.Empty);

            return html.Trim();
        }


        public Table CreateTable(string[][] data)
        {
            return TableCreator.CreateTable(data);

        }

    }
}