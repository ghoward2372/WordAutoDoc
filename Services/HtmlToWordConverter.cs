using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
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

            Console.WriteLine($"=== Converting HTML Content ===\n{html.Trim()}");

            if (html.Contains("[[AcronymTable"))
                return html;

            // First process tables
            var tableMatches = Regex.Matches(html, @"<table[^>]*>(.*?)</table>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            Console.WriteLine($"Found {tableMatches.Count} table(s) in HTML content");

            foreach (Match tableMatch in tableMatches)
            {
                try
                {
                    Console.WriteLine($"=== Processing Table Match ===\n{tableMatch.Value}");
                    var tableData = ExtractTableData(tableMatch.Value);
                    var wordTable = CreateTable(tableData);

                    // Add namespace declaration to ensure proper XML structure
                    var tableXml = wordTable.OuterXml;
                    if (!tableXml.Contains("xmlns:w="))
                    {
                        tableXml = tableXml.Replace("<w:tbl>", $"<w:tbl xmlns:w=\"{WordMlNamespace}\">");
                    }

                    Console.WriteLine($"=== Created Word Table XML ===\n{tableXml}");
                    html = html.Replace(tableMatch.Value, tableXml);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing table: {ex.Message}");
                    continue;
                }
            }

            // Handle other HTML elements
            html = Regex.Replace(html, @"<br\s*/>", "\n");
            html = Regex.Replace(html, @"<p.*?>", string.Empty);
            html = Regex.Replace(html, @"</p>", "\n");
            html = Regex.Replace(html, @"<div.*?>", string.Empty);
            html = Regex.Replace(html, @"</div>", "\n");
            html = Regex.Replace(html, @"<span.*?>", string.Empty);
            html = Regex.Replace(html, @"</span>", string.Empty);

            // Convert HTML entities
            html = WebUtility.HtmlDecode(html);

            // Remove any remaining HTML tags except Word XML
            html = Regex.Replace(html, @"<(?!w:)[^>]+>", string.Empty);

            Console.WriteLine($"=== Final Converted Content ===\n{html.Trim()}");
            return html.Trim();
        }

        private string[][] ExtractTableData(string tableHtml)
        {
            var rows = new List<string[]>();

            // Extract rows
            var rowMatches = Regex.Matches(tableHtml, @"<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            Console.WriteLine($"Found {rowMatches.Count} table rows");

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

        public Table CreateTable(string[][] data)
        {
            if (data == null || data.Length == 0 || data[0].Length == 0)
                throw new ArgumentException("Table data cannot be null or empty");

            Console.WriteLine($"Creating table with {data.Length} rows");
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

                // Add header row properties
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

                    // Add header styling for first row
                    if (rowIndex == 0)
                    {
                        cellProps.AppendChild(new Shading { Fill = "EEEEEE" });
                    }
                    cell.AppendChild(cellProps);

                    // Create run with text
                    var run = new Run();
                    if (rowIndex == 0)
                    {
                        run.AppendChild(new RunProperties(new Bold()));
                    }
                    run.AppendChild(new Text(colIndex < rowData.Length ? rowData[colIndex] : string.Empty));

                    // Create paragraph with run
                    var para = new Paragraph(
                        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                        run
                    );

                    cell.AppendChild(para);
                    row.AppendChild(cell);
                }

                table.AppendChild(row);
            }

            Console.WriteLine($"Created table with {data.Length} rows and {data[0].Length} columns");
            return table;
        }
    }
}