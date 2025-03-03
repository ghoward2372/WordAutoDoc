using DocumentProcessor.Services;
using System.Threading.Tasks;

namespace DocumentProcessor.Models.TagProcessors
{
    public class AcronymTableTagProcessor : ITagProcessor
    {
        private readonly AcronymProcessor _acronymProcessor;
        private readonly IHtmlToWordConverter _htmlConverter;

        public AcronymTableTagProcessor(AcronymProcessor acronymProcessor, IHtmlToWordConverter htmlConverter)
        {
            _acronymProcessor = acronymProcessor;
            _htmlConverter = htmlConverter;
        }

        public Task<ProcessingResult> ProcessTagAsync(string tagContent)
        {
            return ProcessTagAsync(tagContent, null);
        }

        public Task<ProcessingResult> ProcessTagAsync(string tagContent, DocumentProcessingOptions? options)
        {
            var tableData = _acronymProcessor.GetAcronymTableData();
            if (tableData.Length <= 1) // Only header row
                return Task.FromResult(ProcessingResult.FromText("No acronyms found."));

            var table = _htmlConverter.CreateTable(tableData);
            return Task.FromResult(ProcessingResult.FromTable(table));
        }
    }
}