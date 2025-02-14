using System;
using System.Threading.Tasks;
using DocumentProcessor.Services;

namespace DocumentProcessor.Models.TagProcessors
{
    public class WorkItemTagProcessor : ITagProcessor
    {
        private readonly IAzureDevOpsService _azureDevOpsService;
        private readonly IHtmlToWordConverter _htmlConverter;

        public WorkItemTagProcessor(IAzureDevOpsService azureDevOpsService, IHtmlToWordConverter htmlConverter)
        {
            _azureDevOpsService = azureDevOpsService ?? throw new ArgumentNullException(nameof(azureDevOpsService));
            _htmlConverter = htmlConverter ?? throw new ArgumentNullException(nameof(htmlConverter));
        }

        public async Task<string> ProcessTagAsync(string tagContent)
        {
            if (!int.TryParse(tagContent, out int workItemId))
            {
                throw new ArgumentException($"Invalid work item ID: {tagContent}");
            }

            var documentText = await _azureDevOpsService.GetWorkItemDocumentTextAsync(workItemId);
            return _htmlConverter.ConvertHtmlToWordFormat(documentText ?? string.Empty);
        }
    }
}