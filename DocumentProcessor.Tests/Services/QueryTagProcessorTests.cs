using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using DocumentProcessor.Models.TagProcessors;
using DocumentProcessor.Services;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Moq;
using Xunit;
using DocumentFormat.OpenXml.Wordprocessing;

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
            var queryColumns = new List<WorkItemQueryColumn>
            {
                new WorkItemQueryColumn { Name = "ID", ReferenceName = "System.Id" },
                new WorkItemQueryColumn { Name = "Title", ReferenceName = "System.Title" }
            };

            var query = new QueryHierarchyItem
            {
                Columns = queryColumns
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
                        { "System.Title", "First Item" },
                        { "System.State", "Active" }  // This field should not appear in the table
                    }
                },
                new WorkItem
                {
                    Id = 2,
                    Fields = new Dictionary<string, object>
                    {
                        { "System.Id", "2" },
                        { "System.Title", "Second Item" },
                        { "System.State", "Closed" }  // This field should not appear in the table
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
                data => data.Length == expectedTableData.Length &&
                       data[0].Length == expectedTableData[0].Length)))
                .Returns(mockTable);

            // Act
            var result = await _processor.ProcessTagAsync(queryId);

            // Assert
            _mockAzureDevOpsService.Verify(x => x.GetQueryAsync(queryId), Times.Once);
            _mockAzureDevOpsService.Verify(x => x.ExecuteQueryAsync(queryId), Times.Once);
            _mockAzureDevOpsService.Verify(
                x => x.GetWorkItemsAsync(
                    It.Is<IEnumerable<int>>(ids => ids.Contains(1) && ids.Contains(2)),
                    It.Is<IEnumerable<string>>(fields => 
                        fields.Contains("System.Id") && 
                        fields.Contains("System.Title") &&
                        !fields.Contains("System.State"))),  // Verify only requested fields are fetched
                Times.Once);
            _mockHtmlConverter.Verify(x => x.CreateTable(It.IsAny<string[][]>()), Times.Once);
        }

        [Fact]
        public async Task ProcessTagAsync_WithNoColumns_ReturnsErrorMessage()
        {
            // Arrange
            var queryId = Guid.NewGuid().ToString();
            var query = new QueryHierarchyItem
            {
                Columns = new List<WorkItemQueryColumn>()
            };

            _mockAzureDevOpsService.Setup(x => x.GetQueryAsync(queryId)).ReturnsAsync(query);

            // Act
            var result = await _processor.ProcessTagAsync(queryId);

            // Assert
            Assert.Equal("No columns defined in query.", result);
            _mockAzureDevOpsService.Verify(x => x.ExecuteQueryAsync(queryId), Times.Never);
            _mockAzureDevOpsService.Verify(x => x.GetWorkItemsAsync(It.IsAny<IEnumerable<int>>(), It.IsAny<IEnumerable<string>>()), Times.Never);
        }

        [Fact]
        public async Task ProcessTagAsync_WithNoResults_ReturnsErrorMessage()
        {
            // Arrange
            var queryId = Guid.NewGuid().ToString();
            var queryColumns = new List<WorkItemQueryColumn>
            {
                new WorkItemQueryColumn { Name = "ID", ReferenceName = "System.Id" }
            };

            var query = new QueryHierarchyItem
            {
                Columns = queryColumns
            };

            var queryResult = new WorkItemQueryResult
            {
                WorkItems = new List<WorkItemReference>()
            };

            _mockAzureDevOpsService.Setup(x => x.GetQueryAsync(queryId)).ReturnsAsync(query);
            _mockAzureDevOpsService.Setup(x => x.ExecuteQueryAsync(queryId)).ReturnsAsync(queryResult);

            // Act
            var result = await _processor.ProcessTagAsync(queryId);

            // Assert
            Assert.Equal("No results found for query.", result);
            _mockAzureDevOpsService.Verify(x => x.GetWorkItemsAsync(It.IsAny<IEnumerable<int>>(), It.IsAny<IEnumerable<string>>()), Times.Never);
        }
    }
}