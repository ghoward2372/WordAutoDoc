using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Models;
using DocumentProcessor.Models.TagProcessors;
using DocumentProcessor.Utilities;

namespace DocumentProcessor.Services
{
    public class WordDocumentProcessor
    {
        private readonly DocumentProcessingOptions _options;
        private readonly Dictionary<string, ITagProcessor> _tagProcessors;

        public WordDocumentProcessor(DocumentProcessingOptions options)
        {
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _tagProcessors = new Dictionary<string, ITagProcessor>();

            // Always add AcronymTable processor
            _tagProcessors.Add("AcronymTable", new AcronymTableTagProcessor(options.AcronymProcessor, options.HtmlConverter));

            // Only add Azure DevOps related processors if service is available
            if (options.AzureDevOpsService != null)
            {
                _tagProcessors.Add("WorkItem", new WorkItemTagProcessor(options.AzureDevOpsService, options.HtmlConverter));
                _tagProcessors.Add("QueryID", new QueryTagProcessor(options.AzureDevOpsService, options.HtmlConverter));
            }
        }

        public async Task ProcessDocumentAsync()
        {
            try
            {
                Console.WriteLine("\n=== Starting Document Processing ===");
                Console.WriteLine($"Source document: {_options.SourcePath}");
                Console.WriteLine($"Output document: {_options.OutputPath}");

                File.Copy(_options.SourcePath, _options.OutputPath, true);
                Console.WriteLine("Created output document successfully");

                using (WordprocessingDocument targetDoc = WordprocessingDocument.Open(_options.OutputPath, true))
                {
                    var mainPart = targetDoc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                    var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body is missing");

                    await ProcessDocumentContentAsync(body);
                    mainPart.Document.Save();
                    Console.WriteLine("\n=== Document Processing Completed ===");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n=== Document Processing Failed ===");
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        private async Task ProcessDocumentContentAsync(Body body)
        {
            var paragraphsToProcess = body.Elements<Paragraph>().ToList();
            Console.WriteLine($"\nFound {paragraphsToProcess.Count} paragraphs to process");

            foreach (var paragraph in paragraphsToProcess)
            {
                string text = paragraph.InnerText;
                Console.WriteLine($"\nProcessing paragraph: {text}");

                var result = await ProcessTextAsync(text);

                if (result.IsTable)
                {
                    Console.WriteLine("\n=== Inserting Table ===");
                    try
                    {
                        var table = result.TableElement!;
                        var rowCount = table.Elements<TableRow>().Count();
                        var columnCount = table.Elements<TableRow>().FirstOrDefault()?.Elements<TableCell>().Count() ?? 0;

                        Console.WriteLine("Table properties:");
                        Console.WriteLine($"- Rows: {rowCount}");
                        Console.WriteLine($"- Columns: {columnCount}");
                        Console.WriteLine($"- Border size: {table.GetFirstChild<TableProperties>()?.TableBorders?.TopBorder?.Size ?? 0}");
                        Console.WriteLine($"- Width type: {table.GetFirstChild<TableProperties>()?.TableWidth?.Type ?? TableWidthUnitValues.Auto}");

                        paragraph.InsertBeforeSelf(table);
                        paragraph.Remove();
                        Console.WriteLine("Table inserted successfully");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error inserting table: {ex.Message}");
                        Console.WriteLine($"Table XML structure: {result.TableElement!.OuterXml}");
                        throw;
                    }
                }
                else if (text != result.ProcessedText)
                {
                    Console.WriteLine("Updating paragraph text");
                    paragraph.RemoveAllChildren();
                    paragraph.AppendChild(new Run(new Text(result.ProcessedText)));
                }
            }
        }

        private async Task<ProcessingResult> ProcessTextAsync(string text)
        {
            var result = new ProcessingResult { ProcessedText = text };

            foreach (var tagProcessor in _tagProcessors)
            {
                var pattern = RegexPatterns.GetTagPattern(tagProcessor.Key);
                var matches = pattern.Matches(text);

                foreach (Match match in matches)
                {
                    try
                    {
                        Console.WriteLine($"\nProcessing {tagProcessor.Key} tag: {match.Value}");
                        var tagContent = match.Groups[1].Value;
                        var processedContent = await tagProcessor.Value.ProcessTagAsync(tagContent);

                        if (IsTableXml(processedContent))
                        {
                            Console.WriteLine("Table content detected, creating Word table...");
                            result.IsTable = true;
                            result.TableElement = CreateTableFromXml(processedContent);
                            return result;
                        }

                        text = text.Replace(match.Value, processedContent);
                        Console.WriteLine($"Tag processed successfully");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing {tagProcessor.Key} tag: {ex.Message}");
                        text = text.Replace(match.Value, $"[Error processing {tagProcessor.Key} tag]");
                    }
                }
            }

            result.ProcessedText = _options.AcronymProcessor.ProcessText(text);
            return result;
        }

        private bool IsTableXml(string content)
        {
            return content?.StartsWith("<w:tbl") ?? false;
        }

        private Table CreateTableFromXml(string tableXml)
        {
            try
            {
                Console.WriteLine("\n=== Creating Word Table ===");
                var table = new Table();

                // Add enhanced table properties
                var props = new TableProperties(
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
                table.AppendChild(props);

                Console.WriteLine("Parsing table XML content...");
                using (var stringReader = new StringReader(tableXml))
                using (var xmlReader = XmlReader.Create(stringReader))
                {
                    bool isFirstRow = true;
                    while (xmlReader.Read())
                    {
                        if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.LocalName == "tr")
                        {
                            var row = new TableRow();

                            if (isFirstRow)
                            {
                                row.AppendChild(new TableRowProperties(
                                    new TableRowHeight { Val = 400 },
                                    new TableHeader()
                                ));
                            }

                            while (xmlReader.Read() && !(xmlReader.NodeType == XmlNodeType.EndElement && xmlReader.LocalName == "tr"))
                            {
                                if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.LocalName == "tc")
                                {
                                    var cell = new TableCell();
                                    var cellProps = new TableCellProperties(
                                        new TableCellWidth { Type = TableWidthUnitValues.Auto },
                                        new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                                    );

                                    if (isFirstRow)
                                    {
                                        cellProps.AppendChild(new Shading { Fill = "EEEEEE" });
                                    }

                                    cell.AppendChild(cellProps);

                                    string cellContent = xmlReader.ReadInnerXml();
                                    Console.WriteLine($"Adding cell content: {cellContent}");

                                    var paragraph = new Paragraph(
                                        new ParagraphProperties(
                                            new Justification { Val = JustificationValues.Center },
                                            new SpacingBetweenLines { Before = "0", After = "0" }
                                        ),
                                        new Run(
                                            isFirstRow ? new RunProperties(new Bold()) : null,
                                            new Text(cellContent)
                                        )
                                    );

                                    cell.AppendChild(paragraph);
                                    row.AppendChild(cell);
                                }
                            }
                            table.AppendChild(row);
                            isFirstRow = false;
                        }
                    }
                }

                Console.WriteLine("Table created successfully");
                return table;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating table: {ex.Message}");
                Console.WriteLine($"Table XML content: {tableXml}");
                throw new Exception($"Error creating table from XML: {ex.Message}", ex);
            }
        }
    }
}