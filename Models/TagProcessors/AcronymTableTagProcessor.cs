using System.Threading.Tasks;
using System.Linq;
using DocumentProcessor.Services;

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
            var acronyms = _acronymProcessor.GetAcronyms();
            if (!acronyms.Any())
                return Task.FromResult(string.Empty);

            // Create header row
            var tableData = new[]
            {
                new[] { "Acronym", "Definition" }
            }.Concat(
                acronyms.OrderBy(a => a.Key)
                    .Select(a => new[] { a.Key, a.Value })
            ).ToArray();

            var table = _htmlConverter.CreateTable(tableData);
            return Task.FromResult(table.OuterXml);
        }
    }
}