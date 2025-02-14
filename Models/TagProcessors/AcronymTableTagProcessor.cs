using System.Threading.Tasks;
using DocumentProcessor.Services;

namespace DocumentProcessor.Models.TagProcessors
{
    public class AcronymTableTagProcessor : ITagProcessor
    {
        private readonly AcronymProcessor _acronymProcessor;

        public AcronymTableTagProcessor(AcronymProcessor acronymProcessor)
        {
            _acronymProcessor = acronymProcessor;
        }

        public Task<string> ProcessTagAsync(string tagContent)
        {
            return Task.FromResult(_acronymProcessor.GenerateAcronymTable());
        }
    }
}