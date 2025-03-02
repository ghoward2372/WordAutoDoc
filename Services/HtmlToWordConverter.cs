using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Net;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace DocumentProcessor.Services
{
    public class HtmlToWordConverter : IHtmlToWordConverter
    {
        public string ConvertHtmlToWordFormat(string html)
        {
            if (string.IsNullOrEmpty(html))
                return string.Empty;

            // Skip processing if this is our special AcronymTable tag
            if (html.Contains("[[AcronymTable"))
                return html;

            // First check for tables and convert them
            var tableMatches = Regex.Matches(html, @"<table[^>]*>(.*?)</table>", RegexOptions.Singleline);
            foreach (Match tableMatch in tableMatches)
            {
                try
                {
                    var tableData = ExtractTableData(tableMatch.Value);
                    var wordTable = CreateTable(tableData);
                    // Replace the HTML table with Word table XML
                    html = html.Replace(tableMatch.Value, wordTable.OuterXml);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting HTML table: {ex.Message}");
                    // Keep original table text if conversion fails
                    continue;
                }
            }

            // Then handle other HTML elements
            html = Regex.Replace(html, @"<br\s*/>", "\n");
            html = Regex.Replace(html, @"<p.*?>", "");
            html = Regex.Replace(html, @"</p>", "\n");
            html = Regex.Replace(html, @"<div.*?>", "");
            html = Regex.Replace(html, @"</div>", "\n");
            html = Regex.Replace(html, @"<span.*?>", "");
            html = Regex.Replace(html, @"</span>", "");

            // Convert HTML entities
            html = WebUtility.HtmlDecode(html);

            // Remove any remaining HTML tags except our Word table XML
            html = Regex.Replace(html, @"<(?!w:)[^>]+>", string.Empty);

            return html.Trim();
        }

        private string[][] ExtractTableData(string tableHtml)
        {
            var rows = new List<string[]>();

            // Extract rows
            var rowMatches = Regex.Matches(tableHtml, @"<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline);
            foreach (Match rowMatch in rowMatches)
            {
                var cells = new List<string>();

                // Extract cells (both th and td)
                var cellMatches = Regex.Matches(rowMatch.Value, @"<(td|th)[^>]*>(.*?)</(?:td|th)>", RegexOptions.Singleline);
                foreach (Match cellMatch in cellMatches)
                {
                    // Clean cell content
                    var cellContent = cellMatch.Groups[2].Value;
                    cellContent = Regex.Replace(cellContent, @"<[^>]+>", string.Empty); // Remove any nested HTML
                    cellContent = WebUtility.HtmlDecode(cellContent).Trim();
                    cells.Add(cellContent);
                }

                if (cells.Any())
                    rows.Add(cells.ToArray());
            }

            return rows.ToArray();
        }

        public Table CreateTable(string[][] data)
        {
            if (data == null || data.Length == 0)
                throw new ArgumentException("Table data cannot be null or empty");

            var table = new Table();

            // Table properties
            var tableProperties = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 12 },
                    new BottomBorder { Val = BorderValues.Single, Size = 12 },
                    new LeftBorder { Val = BorderValues.Single, Size = 12 },
                    new RightBorder { Val = BorderValues.Single, Size = 12 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 6 }
                ),
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
                new TableLook { Val = "04A0" }
            );
            table.AppendChild(tableProperties);

            // Define TableGrid columns based on the first row
            int columnCount = data[0].Length;
            var grid = new TableGrid();
            for (int i = 0; i < columnCount; i++)
            {
                grid.AppendChild(new GridColumn());
            }
            table.AppendChild(grid);

            // Iterate through rows
            for (int i = 0; i < data.Length; i++)
            {
                var rowData = data[i];
                var row = new TableRow();

                // Ensure correct number of cells in each row
                for (int j = 0; j < columnCount; j++)
                {
                    var cell = new TableCell();
                    var cellProperties = new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto },
                        new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                    );

                    if (i == 0) // Header row styling
                    {
                        cellProperties.AppendChild(new Shading { Fill = "EEEEEE" });
                    }

                    cell.AppendChild(cellProperties);

                    // Create paragraph with proper formatting
                    var paragraph = new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Center },
                            new SpacingBetweenLines { Before = "0", After = "0" }
                        ),
                        new Run(
                            new OpenXmlElement[]
                            {
                                (i == 0) ? new RunProperties(new Bold()) : new RunProperties(),
                                new Text(j < rowData.Length ? rowData[j] : string.Empty)
                            }
                        )
                    );

                    cell.AppendChild(paragraph);
                    row.AppendChild(cell);
                }

                table.AppendChild(row);
            }

            return table;
        }
    }
}