using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Text;

namespace DocumentProcessor.Services
{
    public class HtmlToWordConverter : IHtmlToWordConverter
    {
        private const string WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const int BulletListNumId = 1;
        private const string LIST_START_MARKER = "<LIST_START>";
        private const string LIST_END_MARKER = "<LIST_END>";
        private const string TABLE_START_MARKER = "<TABLE_START>";
        private const string TABLE_END_MARKER = "<TABLE_END>";

        public string ConvertHtmlToWordFormat(string html)
        {
            if (string.IsNullOrEmpty(html))
                return string.Empty;

            Console.WriteLine("\n=== Starting HTML Conversion ===");
            Console.WriteLine($"Input text length: {html?.Length ?? 0}");

            try
            {
                // Check for bullet lists
                if (IsBulletListContent(html))
                {
                    Console.WriteLine("Converting bullet list content");
                    return $"{LIST_START_MARKER}\n{CreateBulletList(html)}\n{LIST_END_MARKER}";
                }

                // If the input is just a table, convert it directly
                if (IsTableContent(html))
                {
                    Console.WriteLine("Converting isolated table content");
                    var tableData = ExtractTableData(html);
                    var wordTable = CreateTable(tableData);
                    return $"{TABLE_START_MARKER}\n{wordTable.OuterXml}\n{TABLE_END_MARKER}";
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

        private bool IsBulletListContent(string html)
        {
            if (string.IsNullOrEmpty(html)) return false;
            var trimmedHtml = html.Trim();
            // More flexible pattern to catch bullet lists from ADO
            var isList = Regex.IsMatch(trimmedHtml, @"<ul[^>]*>.*?</ul>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            Console.WriteLine($"Is bullet list content: {isList}");
            return isList;
        }

        public string CreateBulletList(string html)
        {
            Console.WriteLine($"Creating bullet list from HTML:\n{html}");
            var listItems = ExtractListItems(html);
            var sb = new StringBuilder();

            foreach (var item in listItems)
            {
                var listParagraph = $@"<w:p xmlns:w=""{WordMlNamespace}"">
                    <w:pPr>
                        <w:numPr>
                            <w:ilvl w:val=""0""/>
                            <w:numId w:val=""{BulletListNumId}""/>
                        </w:numPr>
                    </w:pPr>
                    <w:r><w:t>{WebUtility.HtmlEncode(item)}</w:t></w:r>
                </w:p>";
                sb.AppendLine(listParagraph);
            }

            var result = sb.ToString();
            Console.WriteLine($"Generated bullet list XML:\n{result}");
            return result;
        }

        private List<string> ExtractListItems(string html)
        {
            var items = new List<string>();
            var matches = Regex.Matches(html, @"<li[^>]*>(.*?)</li>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            Console.WriteLine($"Found {matches.Count} list items");

            foreach (Match match in matches)
            {
                var content = match.Groups[1].Value;
                content = Regex.Replace(content, @"<[^>]+>", string.Empty);
                content = WebUtility.HtmlDecode(content).Trim();
                if (!string.IsNullOrEmpty(content))
                {
                    items.Add(content);
                    Console.WriteLine($"Added list item: {content}");
                }
            }

            return items;
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
                        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                        run
                    );

                    cell.AppendChild(para);
                    row.AppendChild(cell);
                }

                table.AppendChild(row);
            }

            return table;
        }

        private string[][] ExtractTableData(string html)
        {
            Console.WriteLine($"Extracting data from table HTML...");
            var rows = new List<string[]>();

            // Extract rows
            var rowMatches = Regex.Matches(html, @"<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
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
            return rows.ToArray();
        }
    }
}