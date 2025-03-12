using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Models;
using DocumentProcessor.Models.TagProcessors;
using DocumentProcessor.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace DocumentProcessor.Services
{
    public class WordDocumentProcessor
    {
        private readonly DocumentProcessingOptions _options;
        private readonly Dictionary<string, ITagProcessor> _tagProcessors;
        private const string WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string TABLE_START_MARKER = "<TABLE_START>";
        private const string TABLE_END_MARKER = "<TABLE_END>";
        private const string LIST_START_MARKER = "<LIST_START>";
        private const string LIST_END_MARKER = "<LIST_END>";
        private int adoBulletIndex = -1;



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
                _tagProcessors.Add("QueryAsList", new QueryTagProcessor(options.AzureDevOpsService, options.HtmlConverter));
                _tagProcessors.Add("SBOM", new SBOMTagProcessor(options.AzureDevOpsService, options.HtmlConverter));
                _tagProcessors.Add("ReferenceTable", new ReferenceTableTagProcessor(options.HtmlConverter, options.ReferenceDocProcessor));
                _tagProcessors.Add("GenerateRTM", new RTMTagProcessor(options.AzureDevOpsService));
            }
        }

        public async Task ProcessDocumentAsync()
        {
            try
            {
                Console.WriteLine($"\n=== Starting Document Processing ===");
                Console.WriteLine($"Source: {_options.SourcePath}");
                Console.WriteLine($"Output: {_options.OutputPath}");

                File.Copy(_options.SourcePath, _options.OutputPath, true);

                using (var doc = WordprocessingDocument.Open(_options.OutputPath, true))
                {
                    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                    var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body is missing");

                    await ProcessDocumentContentAsync(body, mainPart, false);
                    mainPart.Document.Save();

                }

                Console.WriteLine("\n=== Document Processing Complete ===");


                await PostProcessDocumentAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing document: {ex.Message}");
                throw;
            }
        }

        private async Task ProcessDocumentContentAsync(Body body, MainDocumentPart mainDocument, bool postProcessing)
        {
            var paragraphsToProcess = body.Elements<Paragraph>().ToList();
            Console.WriteLine($"Processing {paragraphsToProcess.Count} paragraphs");

            foreach (var paragraph in paragraphsToProcess)
            {
                var text = paragraph.InnerText;
                Console.WriteLine($"\nProcessing paragraph text: {text}");

                ProcessingResult processed = null;
                if (postProcessing != false)
                {
                    processed = await PostProcessingProcessText(text);
                }
                else
                {
                    processed = await ProcessTextAsync(text);
                }

                if (processed.IsTable && processed.TableElement != null)
                {
                    try
                    {
                        var table = processed.TableElement;
                        Console.WriteLine("Inserting table into document...");

                        // Add namespace to table if missing
                        if (!table.OuterXml.Contains("xmlns:w="))
                        {
                            Console.WriteLine("Adding namespace to table XML");
                            var newTable = new Table();
                            newTable.InnerXml = table.OuterXml.Replace("<w:tbl>", $"<w:tbl xmlns:w=\"{WordMlNamespace}\">");
                            table = newTable;
                        }

                        // Insert table and remove original paragraph
                        paragraph.InsertBeforeSelf(table);
                        paragraph.Remove();
                        Console.WriteLine("Table inserted successfully");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error inserting table: {ex.Message}");
                        throw;
                    }
                }
                else if (text != processed.ProcessedText)
                {
                    // Check if this is content fromOnce a WorkItem tag that might contain tables
                    if (processed.ProcessedText.Contains(TABLE_START_MARKER) || processed.ProcessedText.Contains(LIST_START_MARKER))
                    {
                        InsertMixedContent(processed.ProcessedText, paragraph, mainDocument);
                    }
                    else
                    {
                        Console.WriteLine("Updating paragraph text");
                        paragraph.RemoveAllChildren();
                        paragraph.AppendChild(new Run(new Text(processed.ProcessedText)));
                    }
                }
            }
        }

        public void InsertMixedContent(string content, Paragraph paragraph, MainDocumentPart mainDocumentPart)
        {
            try
            {
                Console.WriteLine("Inserting mixed content with tables and lists...");

                // Ensure markers are removed before processing
                content = content.Replace(LIST_START_MARKER, "").Replace(LIST_END_MARKER, "");

                // Split content by table and list markers
                var parts = content.Split(new[] { TABLE_START_MARKER, TABLE_END_MARKER }, StringSplitOptions.RemoveEmptyEntries);

                OpenXmlElement currentElement = paragraph;

                foreach (var part in parts)
                {
                    var trimmedPart = part.Trim();
                    if (string.IsNullOrEmpty(trimmedPart)) continue;

                    if (trimmedPart.StartsWith("<w:tbl"))
                    {
                        Console.WriteLine("Processing embedded table");
                        try
                        {
                            var table = new Table();
                            table.InnerXml = trimmedPart;
                            currentElement.InsertAfterSelf(table);
                            currentElement = table;
                            Console.WriteLine("Table inserted successfully");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error inserting table: " + ex.Message);
                        }
                    }
                    else if (trimmedPart.Contains("<w:numPr")) // Proper list detection
                    {
                        Console.WriteLine("Processing embedded list");
                        try
                        {
                            var xmlDoc = new XmlDocument();
                            var xmlNsManager = new XmlNamespaceManager(xmlDoc.NameTable);
                            xmlNsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                            xmlDoc.LoadXml("<root xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" + trimmedPart + "</root>");
                            var listNodes = xmlDoc.SelectNodes("//w:p", xmlNsManager);
                            if (listNodes == null || listNodes.Count == 0)
                                return;

                            // Define the bullet symbol and the left indent (in twentieths of a point, e.g., "720" = 0.5 inches)
                            string bulletSymbol = "• "; // plain bullet with a trailing space
                            string leftIndent = "720";  // adjust as needed to match your document's bullet style

                            foreach (XmlNode node in listNodes)
                            {
                                // Create a new paragraph with a left indent and the bullet symbol prepended.
                                var listParagraph = new Paragraph(
                                    new ParagraphProperties(
                                        new Indentation() { Left = leftIndent }
                                    ),
                                    new Run(new Text(bulletSymbol + node.InnerText.Trim()))
                                );

                                currentElement.InsertAfterSelf(listParagraph);
                                currentElement = listParagraph;
                            }

                            Console.WriteLine("List inserted successfully");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error processing embedded list: " + ex.Message);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Processing text content");
                        try
                        {
                            var newParagraph = new Paragraph(new Run(new Text(trimmedPart)));
                            currentElement.InsertAfterSelf(newParagraph);
                            currentElement = newParagraph;
                            Console.WriteLine("Text paragraph inserted");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error inserting text paragraph: " + ex.Message);
                        }
                    }
                }
                // Remove original paragraph since we've replaced it with new content
                if (paragraph.Parent != null)
                {
                    paragraph.Remove();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Critical error processing mixed content: {ex.Message}");
                throw;
            }
        }

        private async Task<ProcessingResult> ProcessTextAsync(string text)
        {
            var result = ProcessingResult.FromText(text);

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
                        ProcessingResult processedContent;

                        if (tagProcessor.Value is QueryTagProcessor queryTagProcessor)
                        {
                            if (tagProcessor.Key == "QueryAsList")
                            {
                                processedContent = await queryTagProcessor.ProcessQueryAsListAsync(tagContent);
                            }
                            else if (tagProcessor.Key == "QueryID")
                            {
                                processedContent = await queryTagProcessor.ProcessTagAsync(tagContent, _options);
                            }
                            else
                            {
                                processedContent = await tagProcessor.Value.ProcessTagAsync(tagContent, _options);
                            }
                        }
                        else if (tagProcessor.Key != "ReferenceTable")
                        {
                            processedContent = await tagProcessor.Value.ProcessTagAsync(tagContent, _options);
                        }
                        else
                        {
                            /// Reference Table tags don't do anything here....
                            processedContent = new ProcessingResult();
                            processedContent.ProcessedText = "[[ReferenceTable:true]]";
                        }

                        // If the tag processor returned a table, use it directly
                        if (processedContent.IsTable && processedContent.TableElement != null)
                        {
                            Console.WriteLine("Table found in processed content");
                            return processedContent;
                        }

                        // Otherwise, replace the tag with the processed text
                        text = text.Replace(match.Value, processedContent.ProcessedText);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing {tagProcessor.Key} tag: {ex.Message}");
                        text = text.Replace(match.Value, $"[Error processing {tagProcessor.Key} tag: {ex.Message}]");
                    }
                }
            }

            // Process for reference docs
            _options.ReferenceDocProcessor.ProcessText(text);

            // Process acronyms only for non-table content
            result.ProcessedText = _options.AcronymProcessor.ProcessText(text);
            return result;
        }
        public async Task PostProcessDocumentAsync()
        {
            try
            {
                Console.WriteLine($"\n=== Starting Document Post-Processing ===");
                Console.WriteLine($"Source: {_options.SourcePath}");

                // Create a temporary copy of the source document.
                string finalFileOutputName = Path.GetDirectoryName(_options.OutputPath) + "\\" + Path.GetFileNameWithoutExtension(_options.OutputPath) + "_FINAL.docx";
                File.Copy(_options.OutputPath, finalFileOutputName, true);

                // Open the temporary copy for post-processing.
                using (var doc = WordprocessingDocument.Open(finalFileOutputName, true))
                {
                    var postMainPart = doc.MainDocumentPart
                        ?? throw new InvalidOperationException("Main document part is missing");
                    var postBody = postMainPart.Document?.Body
                        ?? throw new InvalidOperationException("Document body is missing");

                    await ProcessDocumentContentAsync(postBody, postMainPart, true);
                    postMainPart.Document.Save();
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();


                // Replace the original file with the updated temporary file.

                Console.WriteLine("\n=== Document Post-Processing Complete ===");

                Console.WriteLine("\n=== Document Generation Complete. Finished File : " + finalFileOutputName);

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing document: {ex.Message}");
                throw;
            }
        }
        private async Task<ProcessingResult> PostProcessingProcessText(string text)
        {
            var result = ProcessingResult.FromText(text);

            foreach (var tagProcessor in _tagProcessors)
            {
                var pattern = RegexPatterns.GetTagPattern(tagProcessor.Key);
                var matches = pattern.Matches(text);

                foreach (Match match in matches)
                {
                    try
                    {
                        Console.WriteLine($"\nPost Processing {tagProcessor.Key} tag: {match.Value}");
                        var tagContent = match.Groups[1].Value;
                        ProcessingResult processedContent;

                        if (tagProcessor.Value is ReferenceTableTagProcessor refTagProcessor)
                        {
                            processedContent = await tagProcessor.Value.ProcessTagAsync(tagContent, _options);


                            // If the tag processor returned a table, use it directly
                            if (processedContent.IsTable && processedContent.TableElement != null)
                            {
                                Console.WriteLine("Table found in processed content");
                                return processedContent;
                            }

                            // Otherwise, replace the tag with the processed text
                            text = text.Replace(match.Value, processedContent.ProcessedText);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing {tagProcessor.Key} tag: {ex.Message}");
                        text = text.Replace(match.Value, $"[Error processing {tagProcessor.Key} tag: {ex.Message}]");
                    }
                }
            }

            // Process for reference docs
            _options.ReferenceDocProcessor.ProcessText(text);

            // Process acronyms only for non-table content
            result.ProcessedText = _options.AcronymProcessor.ProcessText(text);
            return result;

        }
    }
}