using DocumentProcessor.Models;
using DocumentProcessor.Models.TagProcessors;
using DocumentProcessor.Services;
using System;
using System.Threading.Tasks;
using Xunit;
using Moq;
using System.Collections.Generic;
using DocumentProcessor.Models.Configuration;

namespace DocumentProcessor.Tests.TagProcessors
{
    public class WorkItemTagProcessorTests
    {
        private readonly Mock<IAzureDevOpsService> _mockAzureDevOpsService;
        private readonly Mock<IHtmlToWordConverter> _mockHtmlConverter;
        private readonly WorkItemTagProcessor _processor;
        private readonly DocumentProcessingOptions _options;
        private const string TEST_FQ_FIELD = "System.Description";
        private readonly AcronymConfiguration _acronymConfig;

        public WorkItemTagProcessorTests()
        {
            _mockAzureDevOpsService = new Mock<IAzureDevOpsService>();
            _mockHtmlConverter = new Mock<IHtmlToWordConverter>();
            _processor = new WorkItemTagProcessor(_mockAzureDevOpsService.Object, _mockHtmlConverter.Object);

            _acronymConfig = new AcronymConfiguration
            {
                KnownAcronyms = new Dictionary<string, string>
                {
                    { "API", "Application Programming Interface" },
                    { "GUI", "Graphical User Interface" }
                },
                IgnoredAcronyms = new HashSet<string> { "ID", "XML" }
            };

            _options = new DocumentProcessingOptions
            {
                SourcePath = "test.docx",
                OutputPath = "output.docx",
                AzureDevOpsService = _mockAzureDevOpsService.Object,
                AcronymProcessor = new AcronymProcessor(_acronymConfig),
                HtmlConverter = _mockHtmlConverter.Object,
                FQDocumentField = TEST_FQ_FIELD
            };
        }

        [Fact]
        public async Task ProcessTagAsync_ValidWorkItemId_ReturnsProcessedContent()
        {
            // Arrange
            const int workItemId = 1234;
            const string rawContent = "<p>Test content</p>";
            const string processedContent = "Test content";

            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(workItemId, TEST_FQ_FIELD))
                .ReturnsAsync(rawContent);

            _mockHtmlConverter
                .Setup(x => x.ConvertHtmlToWordFormat(rawContent))
                .Returns(processedContent);

            // Act
            var result = await _processor.ProcessTagAsync(workItemId.ToString(), _options);

            // Assert
            Assert.NotNull(result);
            Assert.False(result.IsTable);
            Assert.Equal(processedContent, result.ProcessedText);
            _mockAzureDevOpsService.Verify(x => x.GetWorkItemDocumentTextAsync(workItemId, TEST_FQ_FIELD), Times.Once);
            _mockHtmlConverter.Verify(x => x.ConvertHtmlToWordFormat(rawContent), Times.Once);
        }

        [Fact]
        public async Task ProcessTagAsync_InvalidWorkItemId_ReturnsErrorMessage()
        {
            // Arrange
            const string invalidId = "invalid";

            // Act
            var result = await _processor.ProcessTagAsync(invalidId, _options);

            // Assert
            Assert.NotNull(result);
            Assert.False(result.IsTable);
            Assert.Contains("Invalid work item ID", result.ProcessedText);
        }

        [Fact]
        public async Task ProcessTagAsync_ServiceReturnsNull_ReturnsEmptyResult()
        {
            // Arrange
            const int workItemId = 1234;
            string? nullContent = null;

            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(workItemId, TEST_FQ_FIELD))
                .ReturnsAsync(nullContent);

            // Act
            var result = await _processor.ProcessTagAsync(workItemId.ToString(), _options);

            // Assert
            Assert.NotNull(result);
            Assert.False(result.IsTable);
            Assert.Equal(string.Empty, result.ProcessedText);
        }
    }
}