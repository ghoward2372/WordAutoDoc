using System;
using System.Threading.Tasks;
using DocumentProcessor.Models;
using DocumentProcessor.Services;
using Moq;
using Xunit;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace DocumentProcessor.Tests.Services
{
    public class WordDocumentProcessorTests : IDisposable
    {
        private readonly Mock<IAzureDevOpsService> _mockAzureDevOpsService;
        private readonly Mock<IHtmlToWordConverter> _mockHtmlConverter;
        private readonly AcronymProcessor _acronymProcessor;
        private readonly string _testFilePath;
        private readonly string _outputFilePath;

        public WordDocumentProcessorTests()
        {
            _mockAzureDevOpsService = new Mock<IAzureDevOpsService>();
            _mockHtmlConverter = new Mock<IHtmlToWordConverter>();
            _acronymProcessor = new AcronymProcessor();
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
                HtmlConverter = _mockHtmlConverter.Object
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
                .Setup(x => x.GetWorkItemDocumentTextAsync(1234))
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
                HtmlConverter = _mockHtmlConverter.Object
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

            _mockAzureDevOpsService.Verify(x => x.GetWorkItemDocumentTextAsync(1234), Times.Once);
            _mockHtmlConverter.Verify(x => x.ConvertHtmlToWordFormat("<p>Test work item content</p>"), Times.Once);
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