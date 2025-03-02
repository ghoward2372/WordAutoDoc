using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Models;
using DocumentProcessor.Models.Configuration;
using DocumentProcessor.Services;
using Moq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace DocumentProcessor.Tests.Services
{
    public class WordDocumentProcessorTests : IDisposable
    {
        private readonly Mock<IAzureDevOpsService> _mockAzureDevOpsService;
        private readonly Mock<IHtmlToWordConverter> _mockHtmlConverter;
        private readonly AcronymProcessor _acronymProcessor;
        private readonly string _testFilePath;
        private readonly string _outputFilePath;
        private const string TEST_FQ_FIELD = "System.Description";

        public WordDocumentProcessorTests()
        {
            _mockAzureDevOpsService = new Mock<IAzureDevOpsService>();
            _mockHtmlConverter = new Mock<IHtmlToWordConverter>();
            _acronymProcessor = new AcronymProcessor(new AcronymConfiguration
            {
                KnownAcronyms = new Dictionary<string, string>
                {
                    { "API", "Application Programming Interface" },
                    { "GUI", "Graphical User Interface" }
                },
                IgnoredAcronyms = new HashSet<string> { "ID", "XML" }
            });
            _testFilePath = Path.Combine(Path.GetTempPath(), $"test_input_{Guid.NewGuid()}.docx");
            _outputFilePath = Path.Combine(Path.GetTempPath(), $"test_output_{Guid.NewGuid()}.docx");

            CreateTestDocument();
        }

        private void CreateTestDocument()
        {
            using var doc = WordprocessingDocument.Create(_testFilePath, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            // Add test content with work item tag
            var para = mainPart.Document.Body!.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text("[[WorkItem:1234]]"));
            mainPart.Document.Save();
        }

        [Fact]
        public async Task ProcessDocument_WithMixedContent_HandlesTableAndTextCorrectly()
        {
            // Arrange
            Console.WriteLine("\n=== Starting Mixed Content Test ===");
            var htmlContent = @"
                Text before table
                <table>
                    <tr><th>Header 1</th><th>Header 2</th></tr>
                    <tr><td>Cell 1</td><td>Cell 2</td></tr>
                </table>
                Text after table with an API reference
            ";

            Console.WriteLine($"Test HTML content:\n{htmlContent}");

            // Mock the work item response to return our test HTML
            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(1234, TEST_FQ_FIELD))
                .ReturnsAsync(htmlContent);

            // Create a real HTML converter for this test
            var options = new DocumentProcessingOptions
            {
                SourcePath = _testFilePath,
                OutputPath = _outputFilePath,
                AzureDevOpsService = _mockAzureDevOpsService.Object,
                AcronymProcessor = _acronymProcessor,
                HtmlConverter = new HtmlToWordConverter(), // Use actual converter
                FQDocumentField = TEST_FQ_FIELD
            };

            try
            {
                // Act
                Console.WriteLine("\n=== Processing Document ===");
                var processor = new WordDocumentProcessor(options);
                await processor.ProcessDocumentAsync();

                // Assert
                Assert.True(File.Exists(_outputFilePath), "Output file was not created");

                using (var doc = WordprocessingDocument.Open(_outputFilePath, false))
                {
                    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                    var body = mainPart.Document.Body ?? throw new InvalidOperationException("Document body is missing");

                    Console.WriteLine($"\n=== Document Body XML ===\n{body.InnerXml}");

                    // First verify table is present
                    var tables = body.Descendants<Table>().ToList();
                    Console.WriteLine($"Found {tables.Count} table(s) in document");
                    Assert.Single(tables, "Expected exactly one table");

                    // Then verify paragraphs contain expected text
                    var paragraphs = body.Elements<Paragraph>().ToList();
                    Console.WriteLine($"Found {paragraphs.Count} paragraph(s) in document");
                    foreach (var para in paragraphs)
                    {
                        Console.WriteLine($"Paragraph text: {para.InnerText}");
                    }

                    Assert.Contains(paragraphs, p => p.InnerText.Contains("Text before table"));
                    Assert.Contains(paragraphs, p => p.InnerText.Contains("Text after table"));
                    Assert.Contains(paragraphs, p => p.InnerText.Contains("API"));

                    Console.WriteLine("\n=== Test Complete ===");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n=== Test Failed ===");
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
                throw;
            }
            finally
            {
                // Cleanup test files
                if (File.Exists(_testFilePath))
                    File.Delete(_testFilePath);
                if (File.Exists(_outputFilePath))
                    File.Delete(_outputFilePath);
            }
        }

        public void Dispose()
        {
            try
            {
                if (File.Exists(_testFilePath))
                    File.Delete(_testFilePath);
                if (File.Exists(_outputFilePath))
                    File.Delete(_outputFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during cleanup: {ex.Message}");
            }
        }
    }
}