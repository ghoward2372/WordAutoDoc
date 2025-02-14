using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using DocumentProcessor.Utilities;

namespace DocumentProcessor.Services
{
    public class AcronymProcessor
    {
        private readonly Dictionary<string, string> _acronyms = new Dictionary<string, string>();

        public string ProcessText(string text)
        {
            var acronymPattern = RegexPatterns.AcronymPattern;
            var matches = acronymPattern.Matches(text);

            foreach (Match match in matches)
            {
                string acronym = match.Groups[1].Value;
                if (!_acronyms.ContainsKey(acronym))
                {
                    string definition = ExtractAcronymDefinition(text, match.Index);
                    if (!string.IsNullOrEmpty(definition))
                    {
                        _acronyms[acronym] = definition;
                    }
                }
            }

            return text;
        }

        private string ExtractAcronymDefinition(string text, int acronymPosition)
        {
            // Look for capitalized words before the acronym
            var precedingText = text.Substring(0, acronymPosition);
            var words = precedingText.Split(' ');
            var capitalizedWords = new List<string>();

            for (int i = words.Length - 1; i >= 0; i--)
            {
                if (Regex.IsMatch(words[i], @"^[A-Z]"))
                {
                    capitalizedWords.Insert(0, words[i]);
                }
                else if (capitalizedWords.Count > 0)
                {
                    break;
                }
            }

            return string.Join(" ", capitalizedWords);
        }

        public string GenerateAcronymTable()
        {
            if (_acronyms.Count == 0)
                return string.Empty;

            var table = new System.Text.StringBuilder();
            table.AppendLine("Acronym | Definition");
            table.AppendLine("--------|------------");

            foreach (var acronym in _acronyms.OrderBy(a => a.Key))
            {
                table.AppendLine($"{acronym.Key} | {acronym.Value}");
            }

            return table.ToString();
        }
    }
}