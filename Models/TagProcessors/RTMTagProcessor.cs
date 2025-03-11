using DocumentProcessor.Services;
using System;
using System.Threading.Tasks;

namespace DocumentProcessor.Models.TagProcessors
{
    public class RTMTagProcessor : ITagProcessor
    {
        private IAzureDevOpsService _adoService;


        public RTMTagProcessor(IAzureDevOpsService azureDevOpsService)
        {
            _adoService = azureDevOpsService;

        }


        public Task<ProcessingResult> ProcessTagAsync(string tagContent)
        {
            return ProcessTagAsync(tagContent);
        }

        public Task<ProcessingResult> ProcessTagAsync(string tagContent, DocumentProcessingOptions? options)
        {
            throw new NotImplementedException();
        }
    }
}
