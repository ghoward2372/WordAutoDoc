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
            _testFilePath = $"test_input_{Guid.NewGuid()}.docx";
            _outputFilePath = $"test_output_{Guid.NewGuid()}.docx";

            CreateTestDocument();
        }

        private void CreateTestDocument()
        {
            using var doc = WordprocessingDocument.Create(_testFilePath, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            // Add test content with work item tag
            var para = mainPart.Document.Body!.AppendChild(new Paragraph());
            para.AppendChild(new Run(new Text("[[WorkItem:1234]]")));

            mainPart.Document.Save();
        }

        [Fact]
        public async Task ProcessDocument_WithHtmlTable_CreatesWordTable()
        {
            // Arrange
            var htmlContent = @"
<table>
    <tr><th>Header 1</th><th>Header 2</th></tr>
    <tr><td>Cell 1</td><td>Cell 2</td></tr>
    <tr><td>Cell 3</td><td>Cell 4</td></tr>
</table>";

            Console.WriteLine($"Test HTML content:\n{htmlContent}");

            // Set up Azure DevOps service mock
            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(1234, TEST_FQ_FIELD))
                .ReturnsAsync(htmlContent);

            // Set up document processing options with actual HtmlToWordConverter
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
                var processor = new WordDocumentProcessor(options);
                await processor.ProcessDocumentAsync();

                // Assert
                Assert.True(File.Exists(_outputFilePath));

                using (var doc = WordprocessingDocument.Open(_outputFilePath, false))
                {
                    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                    var body = mainPart.Document.Body ?? throw new InvalidOperationException("Document body is missing");

                    Console.WriteLine($"Document content:\n{body.InnerXml}");

                    // Find tables in the document
                    var tables = body.Descendants<Table>().ToList();
                    Assert.True(tables.Any(), "No tables found in the document");

                    var table = tables.First();
                    Console.WriteLine($"Found table XML:\n{table.OuterXml}");

                    var rows = table.Elements<TableRow>().ToList();
                    Assert.Equal(3, rows.Count); // Header + 2 data rows

                    // Verify header row
                    var headerCells = rows[0].Elements<TableCell>().ToList();
                    Assert.Equal(2, headerCells.Count);
                    Assert.Equal("Header 1", headerCells[0].InnerText.Trim());
                    Assert.Equal("Header 2", headerCells[1].InnerText.Trim());

                    // Verify first data row
                    var firstRowCells = rows[1].Elements<TableCell>().ToList();
                    Assert.Equal("Cell 1", firstRowCells[0].InnerText.Trim());
                    Assert.Equal("Cell 2", firstRowCells[1].InnerText.Trim());

                    // Verify second data row
                    var secondRowCells = rows[2].Elements<TableCell>().ToList();
                    Assert.Equal("Cell 3", secondRowCells[0].InnerText.Trim());
                    Assert.Equal("Cell 4", secondRowCells[1].InnerText.Trim());
                }
            }
            finally
            {
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