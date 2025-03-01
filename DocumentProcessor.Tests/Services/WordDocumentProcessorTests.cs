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
                AcronymProcessor = new AcronymProcessor(),
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
                AcronymProcessor = new AcronymProcessor(),
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = TEST_FQ_FIELD
            });

            // Act
            string result = processor.ExtractTextFromXml(complexXml);

            // Assert
            Assert.Equal("First Second", result);
        }

        [Fact]
        public void ExtractTextFromXml_WithNestedFormatting_ExtractsTextContent()
        {
            // Arrange
            var complexXml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:pPr>
                    <w:rPr>
                        <w:b/>
                        <w:i/>
                        <w:u w:val=""single""/>
                    </w:rPr>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:b/>
                        <w:color w:val=""FF0000""/>
                    </w:rPr>
                    <w:t>Complex</w:t>
                </w:r>
                <w:r>
                    <w:t xml:space=""preserve""> </w:t>
                </w:r>
                <w:r>
                    <w:rPr>
                        <w:i/>
                    </w:rPr>
                    <w:t>Formatting</w:t>
                </w:r>
            </w:p>";
            var processor = new WordDocumentProcessor(new DocumentProcessingOptions
            {
                SourcePath = "test.docx",
                OutputPath = "output.docx",
                AzureDevOpsService = null,
                AcronymProcessor = new AcronymProcessor(),
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = TEST_FQ_FIELD
            });

            // Act
            string result = processor.ExtractTextFromXml(complexXml);

            // Assert
            Assert.Equal("Complex Formatting", result);
        }


        [Fact]
        public void ExtractTextFromXml_WithPreservedWhitespace_NormalizesSpacing()
        {
            // Arrange
            var complexXml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:r>
                    <w:t xml:space=""preserve"">  Multiple    </w:t>
                </w:r>
                <w:r>
                    <w:t>Spaces</w:t>
                </w:r>
                <w:r>
                    <w:t xml:space=""preserve"">  Here  </w:t>
                </w:r>
            </w:p>";
            var processor = new WordDocumentProcessor(new DocumentProcessingOptions
            {
                SourcePath = "test.docx",
                OutputPath = "output.docx",
                AzureDevOpsService = null,
                AcronymProcessor = new AcronymProcessor(),
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = TEST_FQ_FIELD
            });

            // Act
            string result = processor.ExtractTextFromXml(complexXml);

            // Assert
            Assert.Equal("Multiple Spaces Here", result);
        }

        [Fact]
        public void ExtractTextFromXml_WithEmptyElements_FiltersThemOut()
        {
            // Arrange
            var complexXml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:r>
                    <w:t>First</w:t>
                </w:r>
                <w:r>
                    <w:t xml:space=""preserve"">  </w:t>
                </w:r>
                <w:r>
                    <w:t></w:t>
                </w:r>
                <w:r>
                    <w:t>Last</w:t>
                </w:r>
            </w:p>";
            var processor = new WordDocumentProcessor(new DocumentProcessingOptions
            {
                SourcePath = "test.docx",
                OutputPath = "output.docx",
                AzureDevOpsService = null,
                AcronymProcessor = new AcronymProcessor(),
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = TEST_FQ_FIELD
            });

            // Act
            string result = processor.ExtractTextFromXml(complexXml);

            // Assert
            Assert.Equal("First Last", result);
        }

        [Fact]
        public void ExtractTextFromXml_WithMixedContent_ExtractsOnlyTextContent()
        {
            // Arrange
            var complexXml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:r>
                    <w:rPr>
                        <w:b/>
                        <w:i/>
                    </w:rPr>
                    <w:t>Bold</w:t>
                    <w:tab/>
                    <w:t>Italic</w:t>
                </w:r>
                <w:r>
                    <w:br/>
                    <w:t>Next</w:t>
                    <w:drawing/>
                    <w:t>Line</w:t>
                </w:r>
            </w:p>";
            var processor = new WordDocumentProcessor(new DocumentProcessingOptions
            {
                SourcePath = "test.docx",
                OutputPath = "output.docx",
                AzureDevOpsService = null,
                AcronymProcessor = new AcronymProcessor(),
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = TEST_FQ_FIELD
            });

            // Act
            string result = processor.ExtractTextFromXml(complexXml);

            // Assert
            Assert.Equal("Bold Italic Next Line", result);
        }

        [Fact]
        public void ExtractTextFromXml_WithMultipleParagraphs_ExtractsAllText()
        {
            // Arrange
            var complexXml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:r><w:t>First Paragraph</w:t></w:r>
            </w:p>
            <w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:r><w:t>Second Paragraph</w:t></w:r>
            </w:p>";
            var processor = new WordDocumentProcessor(new DocumentProcessingOptions
            {
                SourcePath = "test.docx",
                OutputPath = "output.docx",
                AzureDevOpsService = null,
                AcronymProcessor = new AcronymProcessor(),
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = TEST_FQ_FIELD
            });

            // Act
            string result = processor.ExtractTextFromXml(complexXml);

            // Assert
            Assert.Equal("First Paragraph Second Paragraph", result);
        }

        [Fact]
        public void ExtractTextFromXml_WithSpecialCharacters_HandlesEntities()
        {
            // Arrange
            var complexXml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:r><w:t>Text &amp; Symbols</w:t></w:r>
                <w:r><w:t>Special &lt;characters&gt;</w:t></w:r>
            </w:p>";
            var processor = new WordDocumentProcessor(new DocumentProcessingOptions
            {
                SourcePath = "test.docx",
                OutputPath = "output.docx",
                AzureDevOpsService = null,
                AcronymProcessor = new AcronymProcessor(),
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = TEST_FQ_FIELD
            });

            // Act
            string result = processor.ExtractTextFromXml(complexXml);

            // Assert
            Assert.Equal("Text & Symbols Special <characters>", result);
        }

        [Fact]
        public void ExtractTextFromXml_WithEmptyFormatting_ExtractsTextOnly()
        {
            // Arrange
            var complexXml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                <w:r>
                    <w:rPr><w:b/><w:i/><w:u/></w:rPr>
                    <w:t>Formatted</w:t>
                </w:r>
                <w:r>
                    <w:rPr></w:rPr>
                    <w:t>Text</w:t>
                </w:r>
            </w:p>";
            var processor = new WordDocumentProcessor(new DocumentProcessingOptions
            {
                SourcePath = "test.docx",
                OutputPath = "output.docx",
                AzureDevOpsService = null,
                AcronymProcessor = new AcronymProcessor(),
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = TEST_FQ_FIELD
            });

            // Act
            string result = processor.ExtractTextFromXml(complexXml);

            // Assert
            Assert.Equal("Formatted Text", result);
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