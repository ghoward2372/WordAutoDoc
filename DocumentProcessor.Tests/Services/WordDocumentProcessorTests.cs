using DocumentProcessor.Models;
using DocumentProcessor.Services;
using Moq;
using Xunit;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Threading.Tasks;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentProcessor.Models.Configuration;

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
        private readonly AcronymConfiguration _acronymConfig;

        public WordDocumentProcessorTests()
        {
            _mockAzureDevOpsService = new Mock<IAzureDevOpsService>();
            _mockHtmlConverter = new Mock<IHtmlToWordConverter>();
            _acronymConfig = new AcronymConfiguration
            {
                KnownAcronyms = new Dictionary<string, string>
                {
                    { "API", "Application Programming Interface" },
                    { "GUI", "Graphical User Interface" }
                },
                IgnoredAcronyms = new HashSet<string> { "ID", "XML" }
            };
            _acronymProcessor = new AcronymProcessor(_acronymConfig);
            _testFilePath = $"test_input_{Guid.NewGuid()}.docx";
            _outputFilePath = $"test_output_{Guid.NewGuid()}.docx";

            CreateTestDocument();
        }

        private void CreateTestDocument()
        {
            using var doc = WordprocessingDocument.Create(_testFilePath, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            // Add test content with acronyms
            var para1 = mainPart.Document.Body.AppendChild(new Paragraph());
            para1.AppendChild(new Run(new Text("The API and GUI are important components.")));

            mainPart.Document.Save();
        }

        [Fact]
        public async Task ProcessDocument_WithHtmlTable_CreatesWordTable()
        {
            // Arrange
            var htmlContent = @"<p>Before table</p>
                <table>
                    <tr><th>Header 1</th><th>Header 2</th></tr>
                    <tr><td>Cell 1</td><td>Cell 2</td></tr>
                    <tr><td>Cell 3</td><td>Cell 4</td></tr>
                </table>
                <p>After table</p>";

            Console.WriteLine($"Test HTML content: {htmlContent}");

            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(1234, TEST_FQ_FIELD))
                .ReturnsAsync(htmlContent);

            // Update test document with work item reference
            using (var doc = WordprocessingDocument.Open(_testFilePath, true))
            {
                var para = doc.MainDocumentPart!.Document.Body.AppendChild(new Paragraph());
                para.AppendChild(new Run(new Text("[[WorkItem:1234]]")));
                doc.MainDocumentPart.Document.Save();
            }

            var options = new DocumentProcessingOptions
            {
                SourcePath = _testFilePath,
                OutputPath = _outputFilePath,
                AzureDevOpsService = _mockAzureDevOpsService.Object,
                AcronymProcessor = _acronymProcessor,
                HtmlConverter = new HtmlToWordConverter(), // Use actual converter
                FQDocumentField = TEST_FQ_FIELD
            };

            var processor = new WordDocumentProcessor(options);

            try
            {
                // Act
                await processor.ProcessDocumentAsync();

                // Assert
                Assert.True(File.Exists(_outputFilePath));
                using (var doc = WordprocessingDocument.Open(_outputFilePath, false))
                {
                    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                    var tables = mainPart.Document.Body.Elements<Table>().ToList();
                    Assert.True(tables.Any(), "No tables found in the document");

                    Console.WriteLine($"Found {tables.Count} tables in the document");

                    var table = tables.First();
                    var rows = table.Elements<TableRow>().ToList();
                    Assert.Equal(3, rows.Count); // Header + 2 data rows

                    // Verify header row
                    var headerCells = rows[0].Elements<TableCell>().ToList();
                    Assert.Equal(2, headerCells.Count);

                    var headerText1 = headerCells[0].InnerText;
                    var headerText2 = headerCells[1].InnerText;

                    Console.WriteLine($"Header cell contents: [{headerText1}], [{headerText2}]");

                    Assert.Equal("Header 1", headerText1);
                    Assert.Equal("Header 2", headerText2);

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
                // Cleanup test files
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

            var processor = new WordDocumentProcessor(options);

            try
            {
                // Act
                await processor.ProcessDocumentAsync();

                // Assert
                Assert.True(File.Exists(_outputFilePath));
                using (var doc = WordprocessingDocument.Open(_outputFilePath, false))
                {
                    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                    var text = mainPart.Document.Body.InnerText;

                    Console.WriteLine($"Document text: {text}");

                    // Verify acronyms are present
                    Assert.Contains("API", text);
                    Assert.Contains("GUI", text);
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
            var complexXml = @"<w:tcPr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:tcW w:type=""auto"" /><w:vAlign w:val=""center"" /><w:shd w:fill=""EEEEEE"" /></w:tcPr><w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:pPr><w:jc w:val=""center"" /><w:spacing w:before=""0"" w:after=""0"" /></w:pPr><w:r><w:rPr><w:b /></w:rPr><w:t>ID</w:t></w:r></w:p>";
            var processor = new WordDocumentProcessor(new DocumentProcessingOptions
            {
                SourcePath = "test.docx",
                OutputPath = "output.docx",
                AzureDevOpsService = null,
                AcronymProcessor = new AcronymProcessor(_acronymConfig),
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = TEST_FQ_FIELD
            });

            // Act
            string result = processor.ExtractTextFromXml(complexXml);

            // Assert
            Assert.Equal("ID", result);
        }

        [Fact]
        public void ExtractTextFromXml_WithMultipleTextElements_ExtractsAllText()
        {
            // Arrange
            var complexXml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:r><w:rPr><w:b/></w:rPr><w:t>First</w:t></w:r>
                <w:r><w:t xml:space=""preserve""> </w:t></w:r>
                <w:r><w:rPr><w:i/></w:rPr><w:t>Second</w:t></w:r>
            </w:p>";
            var processor = new WordDocumentProcessor(new DocumentProcessingOptions
            {
                SourcePath = "test.docx",
                OutputPath = "output.docx",
                AzureDevOpsService = null,
                AcronymProcessor = new AcronymProcessor(_acronymConfig),
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = TEST_FQ_FIELD
            });

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