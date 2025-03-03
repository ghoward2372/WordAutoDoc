using System.Threading.Tasks;

namespace DocumentProcessor.Models.TagProcessors
{
    /// <summary>
    /// Interface for processing document tags that can produce either text or table content
    /// </summary>
    public interface ITagProcessor
    {
        /// <summary>
        /// Processes a tag's content and returns either formatted text or a table
        /// </summary>
        /// <param name="tagContent">The content within the tag to process</param>
        /// <returns>A ProcessingResult containing either text or table content</returns>
        Task<ProcessingResult> ProcessTagAsync(string tagContent);

        /// <summary>
        /// Processes a tag's content with additional document processing options
        /// </summary>
        /// <param name="tagContent">The content within the tag to process</param>
        /// <param name="options">Additional processing options</param>
        /// <returns>A ProcessingResult containing either text or table content</returns>
        Task<ProcessingResult> ProcessTagAsync(string tagContent, DocumentProcessingOptions? options);
    }
}