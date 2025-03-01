using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
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
                        cell.AppendChild(new TableRowProperties(new TableHeader())); // Ensure it's a header
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
                        new Text(j < rowData.Length ? rowData[j] : string.Empty) // Ensure no out-of-bounds error
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