using DocumentProcessor.Models;
using System.Threading.Tasks;

namespace DocumentProcessor.Services
{
    public interface ITagProcessor
    {
        Task<ProcessingResult> ProcessTagAsync(string tagContent);
        Task<ProcessingResult> ProcessTagAsync(string tagContent, DocumentProcessingOptions? options);
    }
}
