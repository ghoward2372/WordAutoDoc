using DocumentProcessor.Services;

namespace DocumentProcessor.Models
{
    public class DocumentProcessingOptions
    {
        public required string SourcePath { get; set; }
        public required string OutputPath { get; set; }
        public required AzureDevOpsService AzureDevOpsService { get; set; }
        public required AcronymProcessor AcronymProcessor { get; set; }
        public required HtmlToWordConverter HtmlConverter { get; set; }
    }
}