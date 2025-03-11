using DocumentProcessor.Services;

namespace DocumentProcessor.Models
{
    public class DocumentProcessingOptions
    {
        public required string FQDocumentField { get; set; }
        public required string SourcePath { get; set; }
        public required string OutputPath { get; set; }
        public required IAzureDevOpsService? AzureDevOpsService { get; set; }
        public required AcronymProcessor AcronymProcessor { get; set; }
        public required IHtmlToWordConverter HtmlConverter { get; set; }
        public required ReferenceDocProcessor ReferenceDocProcessor { get; set; }
        public required RTMGenerator RTMGenerator { get; set; }

    }
}