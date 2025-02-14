using System.Text.RegularExpressions;

namespace DocumentProcessor.Utilities
{
    public static class RegexPatterns
    {
        public static readonly Regex AcronymPattern = new Regex(@"\(([A-Z]{2,})\)", RegexOptions.Compiled);

        public static Regex GetTagPattern(string tagType)
        {
            return new Regex($@"\[\[{tagType}:(.+?)\]\]", RegexOptions.Compiled);
        }

        public static readonly Regex CapitalizedWordsPattern = new Regex(@"\b[A-Z][a-z]*\b", RegexOptions.Compiled);
    }
}