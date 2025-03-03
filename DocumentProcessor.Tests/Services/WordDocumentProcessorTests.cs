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
        private const string WORD_ML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public WordDocumentProcessorTests()
        {
            _mockAzureDevOpsService = new Mock<IAzureDevOpsService>();
            _mockHtmlConverter = new Mock<IHtmlToWordConverter>();
            _acronymProcessor = new AcronymProcessor(new AcronymConfiguration
            {
                KnownAcronyms = new Dictionary<string, string>(),
                IgnoredAcronyms = new HashSet<string>()
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

            // Add test content with mixed content (text and table)
            var para = mainPart.Document.Body!.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text("[[WorkItem:1234]]"));
            mainPart.Document.Save();
        }

        [Fact]
        public async Task ProcessDocument_WithMixedContent_HandlesTableAndTextCorrectly()
        {
            // Arrange
            var tableXml = $@"<w:tbl xmlns:w=""{WORD_ML_NAMESPACE}""><w:tblPr><w:tblStyle w:val=""TableGrid""/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Header 1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Header 2</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Cell 1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Cell 2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
            var mixedContent = $@"Converted: Text before table
<TABLE_START>
{tableXml}
<TABLE_END>
Converted: Text after table";

            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(1234, TEST_FQ_FIELD))
                .ReturnsAsync(mixedContent);

            var options = new DocumentProcessingOptions
            {
                SourcePath = _testFilePath,
                OutputPath = _outputFilePath,
                AzureDevOpsService = _mockAzureDevOpsService.Object,
                AcronymProcessor = _acronymProcessor,
                HtmlConverter = _mockHtmlConverter.Object,
                FQDocumentField = TEST_FQ_FIELD
            };

            try
            {
                // Act
                var processor = new WordDocumentProcessor(options);
                await processor.ProcessDocumentAsync();

                // Assert
                using (var doc = WordprocessingDocument.Open(_outputFilePath, false))
                {
                    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                    var body = mainPart.Document.Body ?? throw new InvalidOperationException("Document body is missing");

                    // First verify table is present
                    var tables = body.Descendants<Table>().ToList();
                    Assert.Single(tables, "Expected exactly one table");

                    // Then verify text content
                    var paragraphs = body.Elements<Paragraph>().ToList();
                    Assert.Contains(paragraphs, p => p.InnerText.Contains("Text before table"));
                    Assert.Contains(paragraphs, p => p.InnerText.Contains("Text after table"));

                    // Verify table structure
                    var table = tables.First();
                    var rows = table.Elements<TableRow>().ToList();
                    Assert.Equal(2, rows.Count);
                    Assert.Contains("Header 1", rows[0].InnerText);
                    Assert.Contains("Cell 1", rows[1].InnerText);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n=== Test Failed ===");
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
                throw;
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