using System.Threading.Tasks;

namespace DocumentProcessor.Models.TagProcessors
{
    public interface ITagProcessor
    {
        Task<string> ProcessTagAsync(string tagContent);
    }
}