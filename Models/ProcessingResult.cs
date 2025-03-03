using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Models
{
    public class ProcessingResult
    {
        public string ProcessedText { get; set; } = string.Empty;
        public bool IsTable { get; set; }
        public Table? TableElement { get; set; }

        public static ProcessingResult FromText(string text)
        {
            return new ProcessingResult { ProcessedText = text };
        }

        public static ProcessingResult FromTable(Table table)
        {
            return new ProcessingResult
            {
                IsTable = true,
                TableElement = table
            };
        }
    }
}