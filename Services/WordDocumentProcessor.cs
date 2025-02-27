using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Models;
using DocumentProcessor.Models.TagProcessors;
using DocumentProcessor.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml;

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
                // Create a copy of the source document
                File.Copy(_options.SourcePath, _options.OutputPath, true);

                using (WordprocessingDocument targetDoc = WordprocessingDocument.Open(_options.OutputPath, true))
                {
                    var mainPart = targetDoc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                    var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body is missing");

                    await ProcessDocumentContentAsync(body);
                    mainPart.Document.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error processing document: {ex.Message}", ex);
            }
        }

        private async Task ProcessDocumentContentAsync(Body body)
        {
            var paragraphsToProcess = body.Elements<Paragraph>().ToList();

            foreach (var paragraph in paragraphsToProcess)
            {
                string text = paragraph.InnerText;
                var result = await ProcessTextAsync(text);

                if (result.IsTable)
                {
                    // Insert the table before the current paragraph
                    paragraph.InsertBeforeSelf(result.TableElement!);
                    // Remove the original paragraph that contained the tag
                    paragraph.Remove();
                }
                else if (text != result.ProcessedText)
                {
                    // Replace the paragraph content with processed text
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
                        var tagContent = match.Groups[1].Value;
                        var processedContent = await tagProcessor.Value.ProcessTagAsync(tagContent);

                        if (IsTableXml(processedContent))
                        {
                            result.IsTable = true;
                            result.TableElement = CreateTableFromXml(processedContent);
                            return result;
                        }

                        text = text.Replace(match.Value, processedContent);
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

                // Parse the XML to extract row and cell data
                using (var stringReader = new StringReader(tableXml))
                using (var xmlReader = XmlReader.Create(stringReader))
                {
                    while (xmlReader.Read())
                    {
                        if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.LocalName == "tr")
                        {
                            var row = new TableRow();

                            // Read cells until we reach the end of the row
                            while (xmlReader.Read() && !(xmlReader.NodeType == XmlNodeType.EndElement && xmlReader.LocalName == "tr"))
                            {
                                if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.LocalName == "tc")
                                {
                                    var cell = new TableCell();
                                    var cellProps = new TableCellProperties(
                                        new TableCellWidth { Type = TableWidthUnitValues.Auto }
                                    );
                                    cell.AppendChild(cellProps);

                                    // Get the text content of the cell
                                    string cellContent = xmlReader.ReadInnerXml();
                                    cell.AppendChild(new Paragraph(new Run(new Text(cellContent))));
                                    row.AppendChild(cell);
                                }
                            }
                            table.AppendChild(row);
                        }
                    }
                }

                return table;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating table from XML: {ex.Message}\nXML content: {tableXml}");
                throw new Exception($"Error creating table from XML: {ex.Message}", ex);
            }
        }
    }
}