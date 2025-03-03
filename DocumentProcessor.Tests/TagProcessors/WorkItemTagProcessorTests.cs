using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Models;
using DocumentProcessor.Models.TagProcessors;
using DocumentProcessor.Services;
using DocumentProcessor.Models.Configuration;
using Moq;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;

namespace DocumentProcessor.Tests.TagProcessors
{
    public class WorkItemTagProcessorTests
    {
        private readonly Mock<IAzureDevOpsService> _mockAzureDevOpsService;
        private readonly Mock<IHtmlToWordConverter> _mockHtmlConverter;
        private readonly Mock<ITextBlockProcessor> _mockTextBlockProcessor;
        private readonly WorkItemTagProcessor _processor;
        private readonly DocumentProcessingOptions _options;
        private const string TEST_FQ_FIELD = "System.Description";
        private const string WORD_ML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public WorkItemTagProcessorTests()
        {
            _mockAzureDevOpsService = new Mock<IAzureDevOpsService>();
            _mockHtmlConverter = new Mock<IHtmlToWordConverter>();
            _mockTextBlockProcessor = new Mock<ITextBlockProcessor>();
            _processor = new WorkItemTagProcessor(
                _mockAzureDevOpsService.Object,
                _mockHtmlConverter.Object,
                _mockTextBlockProcessor.Object);

            var acronymConfig = new AcronymConfiguration
            {
                KnownAcronyms = new Dictionary<string, string>(),
                IgnoredAcronyms = new HashSet<string>()
            };

            _options = new DocumentProcessingOptions
            {
                SourcePath = "test.docx",
                OutputPath = "output.docx",
                AzureDevOpsService = _mockAzureDevOpsService.Object,
                HtmlConverter = _mockHtmlConverter.Object,
                AcronymProcessor = new AcronymProcessor(acronymConfig),
                FQDocumentField = TEST_FQ_FIELD
            };
        }

        [Fact]
        public async Task ProcessTagAsync_WithMixedContent_HandlesTableAndTextCorrectly()
        {
            // Arrange
            const int workItemId = 1234;
            var mixedContent = @"
                Text before table
                <table>
                    <tr><th>Header 1</th><th>Header 2</th></tr>
                    <tr><td>Cell 1</td><td>Cell 2</td></tr>
                </table>
                Text after table";

            // Set up text blocks for segmentation
            var textBlocks = new List<TextBlockProcessor.TextBlock>
            {
                new TextBlockProcessor.TextBlock
                {
                    Type = TextBlockProcessor.BlockType.Text,
                    Content = "Text before table"
                },
                new TextBlockProcessor.TextBlock
                {
                    Type = TextBlockProcessor.BlockType.Table,
                    Content = @"<table>
                        <tr><th>Header 1</th><th>Header 2</th></tr>
                        <tr><td>Cell 1</td><td>Cell 2</td></tr>
                    </table>"
                },
                new TextBlockProcessor.TextBlock
                {
                    Type = TextBlockProcessor.BlockType.Text,
                    Content = "Text after table"
                }
            };

            _mockTextBlockProcessor
                .Setup(x => x.SegmentText(mixedContent))
                .Returns(textBlocks);

            var mockTable = new Table();
            var tableXml = $@"<w:tbl xmlns:w=""{WORD_ML_NAMESPACE}""><w:tblPr><w:tblStyle w:val=""TableGrid""/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Header 1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Header 2</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Cell 1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Cell 2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
            mockTable.InnerXml = tableXml;

            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(workItemId, TEST_FQ_FIELD))
                .ReturnsAsync(mixedContent);

            _mockHtmlConverter
                .Setup(x => x.ConvertHtmlToWordFormat(It.Is<string>(s => !s.Contains("<table>"))))
                .Returns<string>(text => "Converted: " + text.Trim());

            _mockHtmlConverter
                .Setup(x => x.CreateTable(It.IsAny<string[][]>()))
                .Returns(mockTable);

            // Act
            var result = await _processor.ProcessTagAsync(workItemId.ToString(), _options);

            // Assert
            Assert.NotNull(result);
            Assert.False(result.IsTable);
            Assert.Contains("<TABLE_START>", result.ProcessedText);
            Assert.Contains("<TABLE_END>", result.ProcessedText);
            Assert.Contains("Converted: Text before table", result.ProcessedText);
            Assert.Contains("Converted: Text after table", result.ProcessedText);
            Assert.Contains(tableXml.Replace(" ", ""), result.ProcessedText.Replace(" ", ""));

            _mockAzureDevOpsService.Verify(x => x.GetWorkItemDocumentTextAsync(workItemId, TEST_FQ_FIELD), Times.Once);
            _mockHtmlConverter.Verify(x => x.CreateTable(It.IsAny<string[][]>()), Times.Once);
            _mockTextBlockProcessor.Verify(x => x.SegmentText(mixedContent), Times.Once);
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
            Assert.Contains("[Invalid work item ID", result.ProcessedText);
        }

        [Fact]
        public async Task ProcessTagAsync_ServiceReturnsNull_ReturnsEmptyResult()
        {
            // Arrange
            const int workItemId = 1234;

            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(workItemId, TEST_FQ_FIELD))
                .ReturnsAsync((string?)null);

            // Act
            var result = await _processor.ProcessTagAsync(workItemId.ToString(), _options);

            // Assert
            Assert.NotNull(result);
            Assert.False(result.IsTable);
            Assert.Contains("[Work Item not found or empty]", result.ProcessedText);
        }

        [Fact]
        public async Task ProcessTagAsync_WithBulletList_HandlesListCorrectly()
        {
            // Arrange
            const int workItemId = 1234;
            var mixedContent = @"
                Text before list
                <ul>
                    <li>First bullet point</li>
                    <li>Second bullet point</li>
                    <li>Third bullet point</li>
                </ul>
                Text after list";

            // Set up text blocks for segmentation
            var textBlocks = new List<TextBlockProcessor.TextBlock>
            {
                new TextBlockProcessor.TextBlock
                {
                    Type = TextBlockProcessor.BlockType.Text,
                    Content = "Text before list"
                },
                new TextBlockProcessor.TextBlock
                {
                    Type = TextBlockProcessor.BlockType.List,
                    Content = @"<ul>
                        <li>First bullet point</li>
                        <li>Second bullet point</li>
                        <li>Third bullet point</li>
                    </ul>"
                },
                new TextBlockProcessor.TextBlock
                {
                    Type = TextBlockProcessor.BlockType.Text,
                    Content = "Text after list"
                }
            };

            _mockTextBlockProcessor
                .Setup(x => x.SegmentText(mixedContent))
                .Returns(textBlocks);

            // Expected bullet list XML
            var listXml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:pPr>
                    <w:numPr>
                        <w:ilvl w:val=""0""/>
                        <w:numId w:val=""1""/>
                    </w:numPr>
                </w:pPr>
                <w:r><w:t>First bullet point</w:t></w:r>
            </w:p>
            <w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:pPr>
                    <w:numPr>
                        <w:ilvl w:val=""0""/>
                        <w:numId w:val=""1""/>
                    </w:numPr>
                </w:pPr>
                <w:r><w:t>Second bullet point</w:t></w:r>
            </w:p>
            <w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:pPr>
                    <w:numPr>
                        <w:ilvl w:val=""0""/>
                        <w:numId w:val=""1""/>
                    </w:numPr>
                </w:pPr>
                <w:r><w:t>Third bullet point</w:t></w:r>
            </w:p>";

            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemDocumentTextAsync(workItemId, TEST_FQ_FIELD))
                .ReturnsAsync(mixedContent);

            _mockHtmlConverter
                .Setup(x => x.ConvertHtmlToWordFormat(It.Is<string>(s => !s.Contains("<ul>"))))
                .Returns<string>(text => "Converted: " + text.Trim());

            _mockHtmlConverter
                .Setup(x => x.CreateBulletList(It.IsAny<string>()))
                .Returns(listXml);

            // Act
            var result = await _processor.ProcessTagAsync(workItemId.ToString(), _options);

            // Assert
            Assert.NotNull(result);
            Assert.False(result.IsTable);
            Assert.Contains("<LIST_START>", result.ProcessedText);
            Assert.Contains("<LIST_END>", result.ProcessedText);
            Assert.Contains("Converted: Text before list", result.ProcessedText);
            Assert.Contains("Converted: Text after list", result.ProcessedText);
            Assert.Contains(listXml.Replace(" ", ""), result.ProcessedText.Replace(" ", ""));

            _mockAzureDevOpsService.Verify(x => x.GetWorkItemDocumentTextAsync(workItemId, TEST_FQ_FIELD), Times.Once);
            _mockHtmlConverter.Verify(x => x.CreateBulletList(It.IsAny<string>()), Times.Once);
            _mockTextBlockProcessor.Verify(x => x.SegmentText(mixedContent), Times.Once);
        }
    }
}