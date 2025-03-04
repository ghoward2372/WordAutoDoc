using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace DocumentProcessor.Services
{
    public class HtmlToWordConverter : IHtmlToWordConverter
    {
        private const string WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public string ConvertHtmlToWordFormat(string html)
        {
            if (string.IsNullOrEmpty(html))
                return string.Empty;

            Console.WriteLine("\n=== Starting HTML Conversion ===");
            Console.WriteLine($"Input text length: {html?.Length ?? 0}");

            if (html.Contains("[[AcronymTable"))
                return html;

            try
            {
                // If the input is just a table, convert it directly
                if (IsTableContent(html))
                {
                    Console.WriteLine("Converting isolated table content");
                    var tableData = ExtractTableData(html);
                    var wordTable = CreateTable(tableData);
                    var tableXml = wordTable.OuterXml;

                    // Ensure proper table XML structure with namespace
                    if (!tableXml.Contains("xmlns:w="))
                    {
                        tableXml = tableXml.Replace("<w:tbl>", $"<w:tbl xmlns:w=\"{WordMlNamespace}\">");
                    }

                    Console.WriteLine($"Generated table XML:\n{tableXml}");
                    return tableXml;
                }

                // Process regular HTML content
                var processedHtml = html;

                // Handle common HTML elements
                processedHtml = Regex.Replace(processedHtml, @"<br\s*/>", "\n");
                processedHtml = Regex.Replace(processedHtml, @"<p.*?>", string.Empty);
                processedHtml = Regex.Replace(processedHtml, @"</p>", "\n");
                processedHtml = Regex.Replace(processedHtml, @"<div.*?>", string.Empty);
                processedHtml = Regex.Replace(processedHtml, @"</div>", "\n");
                processedHtml = Regex.Replace(processedHtml, @"<span.*?>", string.Empty);
                processedHtml = Regex.Replace(processedHtml, @"</span>", string.Empty);

                // Convert HTML entities
                processedHtml = WebUtility.HtmlDecode(processedHtml);

                // Remove any remaining HTML tags except Word XML
                processedHtml = Regex.Replace(processedHtml, @"<(?!w:)[^>]+>", string.Empty);

                Console.WriteLine($"Processed HTML content:\n{processedHtml}");
                return processedHtml.Trim();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting HTML: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        public string ConvertListToWordFormat(string htmlList, int numId)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(htmlList);
            var sb = new StringBuilder();

            void ProcessList(HtmlNode listNode, int level)
            {
                foreach (var listItem in listNode.SelectNodes("li") ?? new HtmlNodeCollection(null))
                {
                    sb.AppendLine("<w:p>");
                    sb.AppendLine($"<w:pPr><w:pStyle w:val=\"ListParagraph\"/><w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"{numId}\"/></w:numPr><w:ind w:left=\"{720 + (level * 360)}\"/></w:pPr>");

                    foreach (var childNode in listItem.ChildNodes)
                    {
                        string textContent = System.Security.SecurityElement.Escape(childNode.InnerText.Trim());
                        if (!string.IsNullOrWhiteSpace(textContent))
                        {
                            sb.AppendLine("<w:r>");
                            sb.AppendLine($"<w:t xml:space=\"preserve\">{textContent}</w:t>");
                            sb.AppendLine("</w:r>");
                        }
                    }
                    sb.AppendLine("</w:p>");

                    var subList = listItem.SelectSingleNode("ul | ol");
                    if (subList != null)
                    {
                        ProcessList(subList, level + 1);
                    }
                }
            }

            var rootList = doc.DocumentNode.SelectSingleNode("ul | ol");
            if (rootList != null)
            {
                ProcessList(rootList, 0);
            }

            string finalXml = sb.ToString();
            Console.WriteLine("Generated Word List XML:\n" + finalXml); // DEBUG OUTPUT
            return finalXml;
        }



        private bool IsTableContent(string html)
        {
            var trimmedHtml = html.Trim();
            var isTable = Regex.IsMatch(trimmedHtml, @"^\s*<table[^>]*>.*?</table>\s*$", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            Console.WriteLine($"Is table content: {isTable}");
            return isTable;
        }

        public Table CreateTable(string[][] data)
        {
            if (data == null || data.Length == 0 || data[0].Length == 0)
                throw new ArgumentException("Table data cannot be null or empty");

            Console.WriteLine($"Creating table with {data.Length} rows and {data[0].Length} columns");

            var table = new Table();

            // Add table properties
            var tableProps = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 12 },
                    new BottomBorder { Val = BorderValues.Single, Size = 12 },
                    new LeftBorder { Val = BorderValues.Single, Size = 12 },
                    new RightBorder { Val = BorderValues.Single, Size = 12 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 6 }
                ),
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }
            );
            table.AppendChild(tableProps);

            // Create grid columns
            var grid = new TableGrid();
            for (int i = 0; i < data[0].Length; i++)
            {
                grid.AppendChild(new GridColumn());
            }
            table.AppendChild(grid);

            // Create rows
            for (int rowIndex = 0; rowIndex < data.Length; rowIndex++)
            {
                var row = new TableRow();
                var rowData = data[rowIndex];

                // Add header properties if this is the first row
                if (rowIndex == 0)
                {
                    row.AppendChild(new TableRowProperties(new TableRowHeight { Val = 400 }));
                }

                // Create cells
                for (int colIndex = 0; colIndex < data[0].Length; colIndex++)
                {
                    var cell = new TableCell();
                    var cellProps = new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto },
                        new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                    );

                    if (rowIndex == 0)
                    {
                        cellProps.AppendChild(new Shading { Fill = "EEEEEE" });
                    }
                    cell.AppendChild(cellProps);

                    var run = new Run();
                    if (rowIndex == 0)
                    {
                        run.AppendChild(new RunProperties(new Bold()));
                    }
                    run.AppendChild(new Text(colIndex < rowData.Length ? rowData[colIndex] : string.Empty));

                    var para = new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = rowIndex == 0 ? JustificationValues.Center : JustificationValues.Left } // Center header, left-align others
                        ),
                        run
                    );

                    cell.AppendChild(para);
                    row.AppendChild(cell);
                }

                table.AppendChild(row);
            }

            return table;
        }

        private string[][] ExtractTableData(string tableHtml)
        {
            Console.WriteLine($"Extracting data from table HTML:\n{tableHtml}");
            var rows = new List<string[]>();

            // Extract rows
            var rowMatches = Regex.Matches(tableHtml, @"<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            Console.WriteLine($"Found {rowMatches.Count} rows");

            foreach (Match rowMatch in rowMatches)
            {
                var cells = new List<string>();

                // Extract cells (both th and td)
                var cellMatches = Regex.Matches(rowMatch.Value, @"<(td|th)[^>]*>(.*?)</(?:td|th)>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                foreach (Match cellMatch in cellMatches)
                {
                    var cellContent = cellMatch.Groups[2].Value;
                    cellContent = Regex.Replace(cellContent, @"<[^>]+>", string.Empty);
                    cellContent = WebUtility.HtmlDecode(cellContent).Trim();
                    cells.Add(cellContent);
                }

                if (cells.Any())
                {
                    rows.Add(cells.ToArray());
                }
            }

            if (!rows.Any())
            {
                throw new InvalidOperationException("No valid data found in table");
            }

            Console.WriteLine($"Extracted {rows.Count} rows with {rows[0].Length} columns");
            foreach (var row in rows)
            {
                Console.WriteLine($"Row data: {string.Join(" | ", row)}");
            }

            return rows.ToArray();
        }
    }
}