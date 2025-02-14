using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Models;
using DocumentProcessor.Models.TagProcessors;
using DocumentProcessor.Utilities;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;

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
            _tagProcessors.Add("AcronymTable", new AcronymTableTagProcessor(options.AcronymProcessor));

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
            foreach (var paragraph in body.Elements<Paragraph>())
            {
                string text = paragraph.InnerText;
                string processedText = await ProcessTextAsync(text);

                if (text != processedText)
                {
                    // Replace the paragraph content with processed text
                    paragraph.RemoveAllChildren();
                    paragraph.AppendChild(new Run(new Text(processedText)));
                }
            }
        }

        private async Task<string> ProcessTextAsync(string text)
        {
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
                        text = text.Replace(match.Value, processedContent);
                    }
                    catch (Exception ex)
                    {
                        // Log the error but continue processing
                        Console.WriteLine($"Error processing {tagProcessor.Key} tag: {ex.Message}");
                        text = text.Replace(match.Value, $"[Error processing {tagProcessor.Key} tag]");
                    }
                }
            }

            text = _options.AcronymProcessor.ProcessText(text);
            return text;
        }
    }
}