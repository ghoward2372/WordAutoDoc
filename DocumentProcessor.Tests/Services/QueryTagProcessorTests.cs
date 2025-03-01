using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Models.TagProcessors;
using DocumentProcessor.Services;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Moq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace DocumentProcessor.Tests.Services
{
    public class QueryTagProcessorTests
    {
        private readonly Mock<IAzureDevOpsService> _mockAzureDevOpsService;
        private readonly Mock<IHtmlToWordConverter> _mockHtmlConverter;
        private readonly QueryTagProcessor _processor;

        public QueryTagProcessorTests()
        {
            _mockAzureDevOpsService = new Mock<IAzureDevOpsService>();
            _mockHtmlConverter = new Mock<IHtmlToWordConverter>();
            _processor = new QueryTagProcessor(_mockAzureDevOpsService.Object, _mockHtmlConverter.Object);
        }

        [Fact]
        public async Task ProcessTagAsync_WithValidQuery_CreatesTableWithQueryResults()
        {
            // Arrange
            var queryId = Guid.NewGuid().ToString();
            var fieldReferences = new List<WorkItemFieldReference>
            {
                new WorkItemFieldReference { Name = "ID", ReferenceName = "System.Id" },
                new WorkItemFieldReference { Name = "Title", ReferenceName = "System.Title" }
            };

            var query = new QueryHierarchyItem
            {
                Columns = fieldReferences
            };

            var queryResult = new WorkItemQueryResult
            {
                WorkItems = new List<WorkItemReference>
                {
                    new WorkItemReference { Id = 1 },
                    new WorkItemReference { Id = 2 }
                }
            };

            var workItems = new List<WorkItem>
            {
                new WorkItem
                {
                    Id = 1,
                    Fields = new Dictionary<string, object>
                    {
                        { "System.Id", "1" },
                        { "System.Title", "First Item" }
                    }
                },
                new WorkItem
                {
                    Id = 2,
                    Fields = new Dictionary<string, object>
                    {
                        { "System.Id", "2" },
                        { "System.Title", "Second Item" }
                    }
                }
            };

            _mockAzureDevOpsService.Setup(x => x.GetQueryAsync(queryId)).ReturnsAsync(query);
            _mockAzureDevOpsService.Setup(x => x.ExecuteQueryAsync(queryId)).ReturnsAsync(queryResult);
            _mockAzureDevOpsService
                .Setup(x => x.GetWorkItemsAsync(
                    It.IsAny<IEnumerable<int>>(),
                    It.IsAny<IEnumerable<string>>()))
                .ReturnsAsync(workItems);

            var expectedTableData = new[]
            {
                new[] { "ID", "Title" },
                new[] { "1", "First Item" },
                new[] { "2", "Second Item" }
            };

            var mockTable = new Table();
            _mockHtmlConverter.Setup(x => x.CreateTable(It.Is<string[][]>(
                data => data.Length == 3 && // Header + 2 data rows
                       data[0].SequenceEqual(new[] { "ID", "Title" }) &&
                       data[1].SequenceEqual(new[] { "1", "First Item" }) &&
                       data[2].SequenceEqual(new[] { "2", "Second Item" })))
                ).Returns(mockTable);

            // Act
            var result = await _processor.ProcessTagAsync(queryId);

            // Assert
            Assert.NotNull(result);
            _mockAzureDevOpsService.Verify(x => x.GetQueryAsync(queryId), Times.Once);
            _mockAzureDevOpsService.Verify(x => x.ExecuteQueryAsync(queryId), Times.Once);
            _mockAzureDevOpsService.Verify(
                x => x.GetWorkItemsAsync(
                    It.Is<IEnumerable<int>>(ids => ids.Count() == 2 &&
                        ids.ToList().All(id => new[] { 1, 2 }.Contains(id))),
                    It.Is<IEnumerable<string>>(fields =>
                        fields.Count() == 2 &&
                        fields.ToList().All(f => new[] { "System.Id", "System.Title" }.Contains(f)))),
                Times.Once);
            _mockHtmlConverter.Verify(x => x.CreateTable(It.Is<string[][]>(
                data => data.Length == 3 && // Header + 2 data rows
                       data[0].SequenceEqual(new[] { "ID", "Title" }) &&
                       data[1].SequenceEqual(new[] { "1", "First Item" }) &&
                       data[2].SequenceEqual(new[] { "2", "Second Item" }))),
                Times.Once);
        }

        [Fact]
        public async Task ProcessTagAsync_WithNoColumns_ReturnsErrorMessage()
        {
            // Arrange
            var queryId = Guid.NewGuid().ToString();
            var query = new QueryHierarchyItem
            {
                Columns = new List<WorkItemFieldReference>()
            };

            _mockAzureDevOpsService.Setup(x => x.GetQueryAsync(queryId)).ReturnsAsync(query);

            // Act
            var result = await _processor.ProcessTagAsync(queryId);

            // Assert
            Assert.Equal("No columns defined in query.", result);
            _mockAzureDevOpsService.Verify(x => x.ExecuteQueryAsync(queryId), Times.Never);
            _mockAzureDevOpsService.Verify(x => x.GetWorkItemsAsync(It.IsAny<IEnumerable<int>>(), It.IsAny<IEnumerable<string>>()), Times.Never);
        }
    }
}