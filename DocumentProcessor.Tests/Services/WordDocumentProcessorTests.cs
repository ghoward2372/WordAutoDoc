using System;
using System.Threading.Tasks;
using DocumentProcessor.Models;
using DocumentProcessor.Services;
using Moq;
using Xunit;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
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
            _testFilePath = "test_input.docx";
            _outputFilePath = "test_output.docx";

            // Create a test document
            TestDocumentGenerator.CreateTestDocument(_testFilePath);
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

            // Act
            await processor.ProcessDocumentAsync();

            // Assert
            Assert.True(File.Exists(_outputFilePath));
            using (var doc = WordprocessingDocument.Open(_outputFilePath, false))
            {
                var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body is missing");
                var text = body.InnerText;

                // Verify that ADO tags are still present (not processed)
                Assert.Contains("[[WorkItem:", text);
                Assert.Contains("[[QueryID:", text);

                // Verify that acronyms were processed
                Assert.Contains("API", text);
                Assert.Contains("GUI", text);
            }
        }

        [Fact]
        public async Task ProcessDocument_WithADO_ProcessesAllTags()
        {
            // Arrange
            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(1234, TEST_FQ_FIELD))
                .ReturnsAsync("<p>Test work item content</p>");

            _mockHtmlConverter
                .Setup(x => x.ConvertHtmlToWordFormat("<p>Test work item content</p>"))
                .Returns("Test work item content");

            var options = new DocumentProcessingOptions
            {
                SourcePath = _testFilePath,
                OutputPath = _outputFilePath,
                AzureDevOpsService = _mockAzureDevOpsService.Object,
                AcronymProcessor = _acronymProcessor,
                HtmlConverter = _mockHtmlConverter.Object,
                FQDocumentField = TEST_FQ_FIELD
            };

            var processor = new WordDocumentProcessor(options);

            // Act
            await processor.ProcessDocumentAsync();

            // Assert
            Assert.True(File.Exists(_outputFilePath));
            using (var doc = WordprocessingDocument.Open(_outputFilePath, false))
            {
                var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body is missing");
                var text = body.InnerText;

                // Verify that work item content was replaced
                Assert.Contains("Test work item content", text);
                Assert.DoesNotContain("[[WorkItem:1234]]", text);
            }

            _mockAzureDevOpsService.Verify(x => x.GetWorkItemDocumentTextAsync(1234, TEST_FQ_FIELD), Times.Once);
            _mockHtmlConverter.Verify(x => x.ConvertHtmlToWordFormat("<p>Test work item content</p>"), Times.Once);
        }

        [Fact]
        public async Task ProcessDocument_WithTableTag_CreatesWordTable()
        {
            // Arrange
            var tableXml = @"<w:tbl xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:tr><w:tc><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl>";

            _mockHtmlConverter
                .Setup(x => x.ConvertHtmlToWordFormat(It.IsAny<string>()))
                .Returns(tableXml);

            var options = new DocumentProcessingOptions
            {
                SourcePath = _testFilePath,
                OutputPath = _outputFilePath,
                AzureDevOpsService = _mockAzureDevOpsService.Object,
                AcronymProcessor = _acronymProcessor,
                HtmlConverter = _mockHtmlConverter.Object,
                FQDocumentField = TEST_FQ_FIELD
            };

            var processor = new WordDocumentProcessor(options);

            // Act
            await processor.ProcessDocumentAsync();

            // Assert
            Assert.True(File.Exists(_outputFilePath));
            using (var doc = WordprocessingDocument.Open(_outputFilePath, false))
            {
                var mainPart = doc.MainDocumentPart;
                var tables = mainPart.Document.Body.Elements<Table>();
                Assert.True(tables.Any());
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
            // Cleanup
            if (File.Exists(_testFilePath))
                File.Delete(_testFilePath);
            if (File.Exists(_outputFilePath))
                File.Delete(_outputFilePath);
        }
    }
}