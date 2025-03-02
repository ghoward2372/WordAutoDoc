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

            // Add test content with both acronyms and work item tag
            var para1 = mainPart.Document.Body!.AppendChild(new Paragraph());
            para1.AppendChild(new Run(new Text("The API and GUI are important components.")));

            var para2 = mainPart.Document.Body.AppendChild(new Paragraph());
            para2.AppendChild(new Run(new Text("[[WorkItem:1234]]")));

            mainPart.Document.Save();
        }

        [Fact]
        public async Task ProcessDocument_WithHtmlTable_CreatesWordTable()
        {
            // Arrange
            var htmlContent = @"<table>
                <tr><th>Header 1</th><th>Header 2</th></tr>
                <tr><td>Cell 1</td><td>Cell 2</td></tr>
                <tr><td>Cell 3</td><td>Cell 4</td></tr>
            </table>";

            Console.WriteLine($"Test HTML content: {htmlContent}");

            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(1234, TEST_FQ_FIELD))
                .ReturnsAsync(htmlContent);

            var options = new DocumentProcessingOptions
            {
                SourcePath = _testFilePath,
                OutputPath = _outputFilePath,
                AzureDevOpsService = _mockAzureDevOpsService.Object,
                AcronymProcessor = _acronymProcessor,
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = TEST_FQ_FIELD
            };

            try
            {
                var processor = new WordDocumentProcessor(options);

                // Act
                await processor.ProcessDocumentAsync();

                // Assert
                Assert.True(File.Exists(_outputFilePath));
                using (var doc = WordprocessingDocument.Open(_outputFilePath, false))
                {
                    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                    var foundTables = mainPart.Document.Body!.Elements<Table>().ToList();

                    // Log document content for debugging
                    Console.WriteLine($"Document content: {mainPart.Document.Body.InnerXml}");
                    foreach (var foundTable in foundTables)
                    {
                        Console.WriteLine($"Found table XML: {foundTable.OuterXml}");
                    }

                    Assert.True(foundTables.Any(), "No tables found in the document");

                    var firstTable = foundTables.First();
                    var rows = firstTable.Elements<TableRow>().ToList();
                    Assert.Equal(3, rows.Count); // Header + 2 data rows

                    // Verify header row
                    var headerCells = rows[0].Elements<TableCell>().ToList();
                    Assert.Equal(2, headerCells.Count);
                    Assert.Equal("Header 1", headerCells[0].InnerText);
                    Assert.Equal("Header 2", headerCells[1].InnerText);

                    // Verify data rows
                    var firstRowCells = rows[1].Elements<TableCell>().ToList();
                    Assert.Equal("Cell 1", firstRowCells[0].InnerText);
                    Assert.Equal("Cell 2", firstRowCells[1].InnerText);

                    var secondRowCells = rows[2].Elements<TableCell>().ToList();
                    Assert.Equal("Cell 3", secondRowCells[0].InnerText);
                    Assert.Equal("Cell 4", secondRowCells[1].InnerText);
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

        [Fact]
        public async Task ProcessDocument_WithoutADO_OnlyProcessesAcronyms()
        {
            // Arrange
            var options = new DocumentProcessingOptions
            {
                SourcePath = _testFilePath,
                OutputPath = _outputFilePath,
                AzureDevOpsService = null,
                AcronymProcessor = _acronymProcessor,
                HtmlConverter = _mockHtmlConverter.Object,
                FQDocumentField = TEST_FQ_FIELD
            };

            try
            {
                var processor = new WordDocumentProcessor(options);

                // Act
                await processor.ProcessDocumentAsync();

                // Assert
                Assert.True(File.Exists(_outputFilePath));
                using (var doc = WordprocessingDocument.Open(_outputFilePath, false))
                {
                    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                    var text = mainPart.Document.Body!.InnerText;

                    Console.WriteLine($"Document text: {text}");

                    // Verify acronyms were processed but ADO tags remain
                    Assert.Contains("API", text);
                    Assert.Contains("GUI", text);
                    Assert.Contains("[[WorkItem:1234]]", text);
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

        [Fact]
        public void ExtractTextFromXml_WithComplexWordXml_ExtractsTextContent()
        {
            // Arrange
            var options = new DocumentProcessingOptions
            {
                SourcePath = "test.docx",
                OutputPath = "output.docx",
                AzureDevOpsService = null,
                AcronymProcessor = _acronymProcessor,
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = TEST_FQ_FIELD
            };

            var processor = new WordDocumentProcessor(options);
            var complexXml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:r><w:rPr><w:b/></w:rPr><w:t>First</w:t></w:r>
                <w:r><w:t xml:space=""preserve""> </w:t></w:r>
                <w:r><w:rPr><w:i/></w:rPr><w:t>Second</w:t></w:r>
            </w:p>";

            // Act
            string result = processor.ExtractTextFromXml(complexXml);

            // Assert
            Assert.Equal("First Second", result);
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