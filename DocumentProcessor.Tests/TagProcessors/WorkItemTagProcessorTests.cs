using System;
using System.Threading.Tasks;
using DocumentProcessor.Models.TagProcessors;
using DocumentProcessor.Services;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Moq;
using Xunit;

namespace DocumentProcessor.Tests.TagProcessors
{
    public class WorkItemTagProcessorTests
    {
        private readonly Mock<IAzureDevOpsService> _mockAzureDevOpsService;
        private readonly Mock<IHtmlToWordConverter> _mockHtmlConverter;
        private readonly WorkItemTagProcessor _processor;

        public WorkItemTagProcessorTests()
        {
            _mockAzureDevOpsService = new Mock<IAzureDevOpsService>();
            _mockHtmlConverter = new Mock<IHtmlToWordConverter>();
            _processor = new WorkItemTagProcessor(_mockAzureDevOpsService.Object, _mockHtmlConverter.Object);
        }

        [Fact]
        public async Task ProcessTagAsync_ValidWorkItemId_ReturnsProcessedContent()
        {
            // Arrange
            const int workItemId = 1234;
            const string rawContent = "<p>Test content</p>";
            const string processedContent = "Test content";

            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(workItemId))
                .ReturnsAsync(rawContent);

            _mockHtmlConverter
                .Setup(x => x.ConvertHtmlToWordFormat(rawContent))
                .Returns(processedContent);

            // Act
            var result = await _processor.ProcessTagAsync(workItemId.ToString());

            // Assert
            Assert.Equal(processedContent, result);
            _mockAzureDevOpsService.Verify(x => x.GetWorkItemDocumentTextAsync(workItemId), Times.Once);
            _mockHtmlConverter.Verify(x => x.ConvertHtmlToWordFormat(rawContent), Times.Once);
        }

        [Fact]
        public async Task ProcessTagAsync_InvalidWorkItemId_ThrowsArgumentException()
        {
            // Arrange
            const string invalidId = "invalid";

            // Act & Assert
            await Assert.ThrowsAsync<ArgumentException>(() => _processor.ProcessTagAsync(invalidId));
        }

        [Fact]
        public async Task ProcessTagAsync_ServiceReturnsNull_ReturnsEmptyString()
        {
            // Arrange
            const int workItemId = 1234;
            string? nullContent = null;

            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(workItemId))
                .ReturnsAsync(nullContent);

            _mockHtmlConverter
                .Setup(x => x.ConvertHtmlToWordFormat(It.IsAny<string>()))
                .Returns(string.Empty);

            // Act
            var result = await _processor.ProcessTagAsync(workItemId.ToString());

            // Assert
            Assert.Equal(string.Empty, result);
        }
    }
}