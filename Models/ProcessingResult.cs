using System;

namespace DocumentProcessor.Models
{
    public class ProcessingResult
    {
        public bool Success { get; set; }
        public string? Message { get; set; }
        public string? ProcessedContent { get; set; }

        public ProcessingResult()
        {
            Success = false;
            Message = null;
            ProcessedContent = null;
        }
    }
}