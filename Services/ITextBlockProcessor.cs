using System.Collections.Generic;

namespace DocumentProcessor.Services
{
    public interface ITextBlockProcessor
    {
        List<TextBlockProcessor.TextBlock> SegmentText(string inputText);
    }
}
