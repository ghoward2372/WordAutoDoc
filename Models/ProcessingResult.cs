using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Models
{
    public class ProcessingResult
    {
        public string ProcessedText { get; set; } = string.Empty;
        public bool IsTable { get; set; }
        public Table? TableElement { get; set; }
    }
}