using System;
using System.Text.RegularExpressions;
using System.Net;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

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
            if (data == null || data.Length == 0)
                return new Table();

            var table = new Table();
            var props = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single },
                    new BottomBorder { Val = BorderValues.Single },
                    new LeftBorder { Val = BorderValues.Single },
                    new RightBorder { Val = BorderValues.Single },
                    new InsideHorizontalBorder { Val = BorderValues.Single },
                    new InsideVerticalBorder { Val = BorderValues.Single }
                )
            );
            table.AppendChild(props);

            for (int i = 0; i < data.Length; i++)
            {
                var tr = new TableRow();
                for (int j = 0; j < data[i].Length; j++)
                {
                    var tc = new TableCell(new Paragraph(new Run(new Text(data[i][j] ?? string.Empty))));
                    tr.AppendChild(tc);
                }
                table.AppendChild(tr);
            }

            return table;
        }
    }
}