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

        public Task<string> ProcessTagAsync(string tagContent)
        {
            var tableData = _acronymProcessor.GetAcronymTableData();
            if (tableData.Length <= 1) // Only header row
                return Task.FromResult("No acronyms found.");

            var table = _htmlConverter.CreateTable(tableData);
            return Task.FromResult(table.OuterXml);
        }

        public Task<string> ProcessTagAsync(string tagContent, DocumentProcessingOptions options)
        {
            return ProcessTagAsync(tagContent);
        }
    }
}