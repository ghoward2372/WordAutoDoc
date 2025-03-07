using DocumentProcessor.Services;
using System.Threading.Tasks;

namespace DocumentProcessor.Models.TagProcessors
{
    public class ReferenceTableTagProcessor : ITagProcessor
    {
        private readonly IHtmlToWordConverter _htmlConverter;
        private readonly ReferenceDocProcessor _refDocProcessor;


        public ReferenceTableTagProcessor(IHtmlToWordConverter htmlConverter, ReferenceDocProcessor refDocProcessor)
        {
            _htmlConverter = htmlConverter;
            _refDocProcessor = refDocProcessor;
        }

        public Task<ProcessingResult> ProcessTagAsync(string tagContent)
        {
            return ProcessTagAsync(tagContent, null);
        }

        public Task<ProcessingResult> ProcessTagAsync(string tagContent, DocumentProcessingOptions? options)
        {
            var tableData = _refDocProcessor.GetReferenceTableData();
            if (tableData.Length <= 1) // Only header row
                return Task.FromResult(ProcessingResult.FromText("No references found."));
            var table = _htmlConverter.CreateTable(tableData);
            return Task.FromResult(ProcessingResult.FromTable(table));
        }
    }
}
