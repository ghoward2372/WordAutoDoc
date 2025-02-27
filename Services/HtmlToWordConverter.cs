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

            // Enhanced table properties for better visual appearance
            var tableProperties = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 12 },
                    new BottomBorder { Val = BorderValues.Single, Size = 12 },
                    new LeftBorder { Val = BorderValues.Single, Size = 12 },
                    new RightBorder { Val = BorderValues.Single, Size = 12 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 6 }
                ),
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }, // 50% of page width
                new TableLook { Val = "04A0" } // Enable header row and banded rows
            );
            table.AppendChild(tableProperties);

            // Add the table grid with column definitions
            var grid = new TableGrid();
            for (int i = 0; i < (data.Length > 0 ? data[0].Length : 0); i++)
            {
                grid.AppendChild(new GridColumn());
            }
            table.AppendChild(grid);

            bool isHeader = true;
            foreach (var rowData in data)
            {
                var row = new TableRow();

                if (isHeader)
                {
                    // Style the header row
                    row.AppendChild(new TableRowProperties(
                        new TableRowHeight { Val = 400 }, // Slightly taller header row
                        new TableHeader() // Mark as header row
                    ));
                }

                foreach (var cellData in rowData)
                {
                    var cell = new TableCell();

                    // Style each cell
                    var cellProperties = new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto },
                        new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                    );

                    if (isHeader)
                    {
                        // Add bold text and center alignment for header cells
                        cellProperties.AppendChild(new Shading { Fill = "EEEEEE" }); // Light gray background
                    }

                    cell.AppendChild(cellProperties);

                    // Create paragraph with proper formatting
                    var paragraph = new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Center },
                            new SpacingBetweenLines { Before = "0", After = "0" }
                        ),
                        new Run(
                            isHeader ? new RunProperties(new Bold()) : null,
                            new Text(cellData ?? string.Empty)
                        )
                    );

                    cell.AppendChild(paragraph);
                    row.AppendChild(cell);
                }

                table.AppendChild(row);
                isHeader = false;
            }

            return table;
        }
    }
}