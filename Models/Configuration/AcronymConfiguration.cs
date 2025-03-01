using System.Collections.Generic;

namespace DocumentProcessor.Models.Configuration
{
    public class AcronymConfiguration
    {
        public required Dictionary<string, string> KnownAcronyms { get; set; }
        public required HashSet<string> IgnoredAcronyms { get; set; }
    }
}
