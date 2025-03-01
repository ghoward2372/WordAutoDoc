using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using DocumentProcessor.Utilities;
using DocumentProcessor.Models.Configuration;

namespace DocumentProcessor.Services
{
    public class AcronymProcessor
    {
        private readonly Dictionary<string, string> _acronyms = new();
        private readonly AcronymConfiguration _configuration;

        public AcronymProcessor(AcronymConfiguration configuration)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));

            // Initialize with known acronyms from configuration
            foreach (var knownAcronym in _configuration.KnownAcronyms)
            {
                _acronyms[knownAcronym.Key] = knownAcronym.Value;
            }
        }

        public string ProcessText(string text)
        {
            var acronymPattern = RegexPatterns.AcronymPattern;
            var matches = acronymPattern.Matches(text);

            foreach (Match match in matches)
            {
                string acronym = match.Groups[1].Value;

                // Skip ignored acronyms
                if (_configuration.IgnoredAcronyms.Contains(acronym))
                {
                    Console.WriteLine($"Skipping ignored acronym: {acronym}");
                    continue;
                }

                if (!_acronyms.ContainsKey(acronym))
                {
                    // Look for definition in the text first
                    string definition = ExtractAcronymDefinition(text, match.Index);
                    if (!string.IsNullOrEmpty(definition))
                    {
                        Console.WriteLine($"Found definition in document for {acronym}: {definition}");
                        _acronyms[acronym] = definition;
                    }
                    else if (_configuration.KnownAcronyms.TryGetValue(acronym, out var knownDefinition))
                    {
                        Console.WriteLine($"Using known definition for {acronym}: {knownDefinition}");
                        _acronyms[acronym] = knownDefinition;
                    }
                    else
                    {
                        Console.WriteLine($"No definition found for acronym: {acronym}");
                        _acronyms[acronym] = string.Empty; // Store empty definition
                    }
                }
            }

            return text;
        }

        private string ExtractAcronymDefinition(string text, int acronymPosition)
        {
            try
            {
                // Look for capitalized words before the acronym
                var precedingText = text.Substring(0, acronymPosition);
                var words = precedingText.Split(' ');
                var capitalizedWords = new List<string>();

                // Work backwards from the acronym position
                for (int i = words.Length - 1; i >= 0; i--)
                {
                    var word = words[i].Trim();
                    if (string.IsNullOrEmpty(word)) continue;

                    // Check if word starts with capital letter
                    if (Regex.IsMatch(word, @"^[A-Z]"))
                    {
                        capitalizedWords.Insert(0, word);
                    }
                    else if (capitalizedWords.Count > 0)
                    {
                        // Stop when we hit a non-capitalized word after finding some capitalized words
                        break;
                    }
                }

                if (capitalizedWords.Count > 0)
                {
                    var definition = string.Join(" ", capitalizedWords);
                    Console.WriteLine($"Extracted definition from text: {definition}");
                    return definition;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting acronym definition: {ex.Message}");
            }

            return string.Empty;
        }

        public Dictionary<string, string> GetAcronyms()
        {
            return new Dictionary<string, string>(_acronyms);
        }

        public string[][] GetAcronymTableData()
        {
            // Create header row
            var tableData = new List<string[]>
            {
                new[] { "Acronym", "Definition" }
            };

            // Add acronym rows, sorted alphabetically
            tableData.AddRange(
                _acronyms
                    .Where(a => !_configuration.IgnoredAcronyms.Contains(a.Key))
                    .OrderBy(a => a.Key)
                    .Select(a => new[] { a.Key, a.Value })
            );

            return tableData.ToArray();
        }
    }
}