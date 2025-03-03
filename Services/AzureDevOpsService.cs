using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DocumentProcessor.Services
{
    public interface IAzureDevOpsService
    {
        Task<string> GetWorkItemDocumentTextAsync(int workItemId, string fqDocumentField);
        Task<WorkItemQueryResult> ExecuteQueryAsync(string queryId);
        Task<QueryHierarchyItem> GetQueryAsync(string queryId);
        Task<IEnumerable<WorkItem>> GetWorkItemsAsync(IEnumerable<int> workItemIds, IEnumerable<string>? fields = null);
    }

    public class AzureDevOpsService : IAzureDevOpsService
    {
        private readonly WorkItemTrackingHttpClient _witClient;
        private static string _projectName;

        public AzureDevOpsService(WorkItemTrackingHttpClient witClient)
        {
            _witClient = witClient ?? throw new ArgumentNullException(nameof(witClient));
        }

        public static AzureDevOpsService Initialize()
        {
            var config = ConfigurationService.LoadAzureDevOpsConfig();
            var credentials = new VssBasicCredential(string.Empty, config.PersonalAccessToken);
            var connection = new VssConnection(new Uri(config.BaseUrl), credentials);
            _projectName = config.ProjectName;

            try
            {
                var witClient = connection.GetClient<WorkItemTrackingHttpClient>();
                return new AzureDevOpsService(witClient);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to initialize Azure DevOps connection. Please verify your organization name and PAT. Details: {ex.Message}");
            }
        }


        public async Task<string> GetWorkItemDocumentTextAsync(int workItemId, string fqDocumentField)
        {
            try
            {
                var workItem = await _witClient.GetWorkItemAsync(workItemId, expand: WorkItemExpand.All);
                if (workItem?.Fields == null)
                {
                    throw new InvalidOperationException($"Work item {workItemId} or its fields are null");
                }

                if (!workItem.Fields.TryGetValue(fqDocumentField, out object? value) || value == null)
                {
                    return string.Empty;
                }

                string htmlContent = value.ToString() ?? string.Empty;
                return ConvertHtmlTablesToWordTables(htmlContent);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving work item {workItemId}: {ex.Message}", ex);
            }
        }

        private string ConvertHtmlTablesToWordTables(string html)
        {
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);

            var tables = doc.DocumentNode.SelectNodes("//table");
            if (tables == null) return html; // No tables found

            foreach (var table in tables)
            {
                string[][] extractedTable = ExtractTableData(table);

                // **NEW CHECK: Skip tables with no data**
                if (extractedTable.Length == 0 || extractedTable.All(row => row == null || row.All(cell => string.IsNullOrWhiteSpace(cell))))
                {
                    continue; // Skip empty or invalid tables
                }

                string wordTableXml = CreateStyledTable(extractedTable).OuterXml;

                html = html.Replace(table.OuterHtml, wordTableXml); // Replace HTML table with Word table XML
            }

            return html;
        }

        private string[][] ExtractTableData(HtmlNode table)
        {
            var rows = table.SelectNodes(".//tr");
            if (rows == null || rows.Count == 0) return new string[0][]; // No rows, return empty array

            List<string[]> tableData = new List<string[]>();

            foreach (var row in rows)
            {
                var cells = row.SelectNodes(".//td | .//th"); // Extracts both headers and cells
                if (cells == null || cells.Count == 0) continue; // **NEW CHECK: Skip empty rows**

                string[] rowData = cells
                    .Select(cell => cell.InnerText.Trim())
                    .Where(cellText => !string.IsNullOrWhiteSpace(cellText)) // **NEW CHECK: Skip empty columns**
                    .ToArray();

                if (rowData.Length > 0) // **Ensure row has at least one valid column**
                {
                    tableData.Add(rowData);
                }
            }

            return tableData.Count > 0 ? tableData.ToArray() : new string[0][]; // **Ensure valid table structure**
        }

        private Table CreateStyledTable(string[][] data)
        {
            if (data == null || data.Length == 0)
                throw new ArgumentException("Table data cannot be null or empty");

            var table = new Table();

            // Table Properties (borders, width, design)
            var tableProperties = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Thick, Size = 24, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Thick, Size = 24, Color = "000000" },
                    new LeftBorder { Val = BorderValues.Thick, Size = 24, Color = "000000" },
                    new RightBorder { Val = BorderValues.Thick, Size = 24, Color = "000000" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 12, Color = "000000" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 12, Color = "000000" }
                ),
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }, // 50% width of page
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

                    // **Enforce Borders for every single cell**
                    var cellBorders = new TableCellBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 12, Color = "000000" },
                        new BottomBorder { Val = BorderValues.Single, Size = 12, Color = "000000" },
                        new LeftBorder { Val = BorderValues.Single, Size = 12, Color = "000000" },
                        new RightBorder { Val = BorderValues.Single, Size = 12, Color = "000000" }
                    );

                    // **Force background color and alignment**
                    TableCellProperties cellProperties = new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = (100 / columnCount).ToString() },
                        new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center },
                        cellBorders
                    );

                    // **Header Row Styling**
                    if (i == 0)
                    {
                        cellProperties.AppendChild(new Shading()
                        {
                            Val = ShadingPatternValues.Clear,
                            Fill = "AAAAAA" // Gray Header Background (more Word-friendly)
                        });
                    }

                    cell.AppendChild(cellProperties);

                    // **Force paragraph formatting**
                    ParagraphProperties paraProperties = new ParagraphProperties(
                        new Justification { Val = JustificationValues.Center }
                    );

                    RunProperties textRunProperties = new RunProperties();
                    if (i == 0) // **Make header bold**
                    {
                        textRunProperties.AppendChild(new Bold());
                    }

                    // **Ensure text rendering properly**
                    Run textRun = new Run(textRunProperties, new Text(j < rowData.Length ? rowData[j] : string.Empty));
                    Paragraph paragraph = new Paragraph(paraProperties, textRun);

                    cell.AppendChild(paragraph);
                    row.AppendChild(cell);
                }

                table.AppendChild(row);
            }

            return table;
        }



        public async Task<WorkItemQueryResult> ExecuteQueryAsync(string queryId)
        {
            try
            {
                if (!Guid.TryParse(queryId, out _))
                    throw new ArgumentException("Invalid query ID format. Expected a GUID.");

                return await _witClient.QueryByIdAsync(new Guid(queryId));
            }
            catch (Exception ex)
            {
                throw new Exception($"Error executing query {queryId}: {ex.Message}", ex);
            }
        }

        public async Task<QueryHierarchyItem> GetQueryAsync(string queryId)
        {
            try
            {
                if (!Guid.TryParse(queryId, out var guid))
                    throw new ArgumentException("Invalid query ID format. Expected a GUID.");

                // The project parameter is required but can be empty for organization-wide queries
                return await _witClient.GetQueryAsync(_projectName, guid.ToString(), QueryExpand.All);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving query {queryId}: {ex.Message}", ex);
            }
        }

        public async Task<IEnumerable<WorkItem>> GetWorkItemsAsync(IEnumerable<int> workItemIds, IEnumerable<string>? fields = null)
        {
            try
            {
                if (!workItemIds.Any())
                    return new List<WorkItem>();

                return await _witClient.GetWorkItemsAsync(
                    workItemIds,
                    fields,
                    expand: WorkItemExpand.None
                );
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving work items: {ex.Message}", ex);
            }
        }
    }
}