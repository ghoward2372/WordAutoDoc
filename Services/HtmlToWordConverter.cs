using System;
using System.Text.RegularExpressions;
using System.Net;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

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
                throw new ArgumentException("Table data cannot be null or empty");

            var table = new Table();
            var props = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single },
                    new BottomBorder { Val = BorderValues.Single },
                    new LeftBorder { Val = BorderValues.Single },
                    new RightBorder { Val = BorderValues.Single },
                    new InsideHorizontalBorder { Val = BorderValues.Single },
                    new InsideVerticalBorder { Val = BorderValues.Single }
                ),
                new TableWidth { Type = TableWidthUnitValues.Auto }
            );
            table.AppendChild(props);

            foreach (var rowData in data)
            {
                var row = new TableRow();
                foreach (var cellData in rowData)
                {
                    var cell = new TableCell();
                    var cellProps = new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto }
                    );
                    cell.AppendChild(cellProps);
                    cell.AppendChild(new Paragraph(new Run(new Text(cellData ?? string.Empty))));
                    row.AppendChild(cell);
                }
                table.AppendChild(row);
            }

            return table;
        }
    }
}