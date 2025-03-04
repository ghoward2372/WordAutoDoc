using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocumentProcessor.Services
{
    public class TextBlockProcessor : ITextBlockProcessor
    {
        public enum BlockType
        {
            Text,
            List,
            Table
        }

        public class TextBlock
        {
            public BlockType Type { get; set; }
            public string Content { get; set; } = string.Empty;
        }

        public List<TextBlock> SegmentText(string inputText)
        {
            Console.WriteLine("=== Starting Text Segmentation ===");
            Console.WriteLine($"Input text length: {inputText?.Length ?? 0}");

            if (string.IsNullOrEmpty(inputText))
            {
                Console.WriteLine("Warning: Input text is null or empty");
                return new List<TextBlock>();
            }

            var blocks = new List<TextBlock>();
            var currentPosition = 0;

            // Find all table and list matches in the text
            var tableMatches = Regex.Matches(inputText, @"<table[^>]*>.*?</table>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            var listMatches = Regex.Matches(inputText, @"<(ul|ol)[^>]*>.*?</\1>", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            Console.WriteLine($"Found {tableMatches.Count} table(s) and {listMatches.Count} list(s) in text");

            // Merge matches and process sequentially
            var matches = tableMatches.Cast<Match>().Concat(listMatches.Cast<Match>()).OrderBy(m => m.Index);

            foreach (var match in matches)
            {
                Console.WriteLine($"Processing match at position {match.Index}, length {match.Length}");

                // Add text before the match if any
                if (match.Index > currentPosition)
                {
                    var textContent = inputText.Substring(currentPosition, match.Index - currentPosition).Trim();
                    if (!string.IsNullOrWhiteSpace(textContent))
                    {
                        Console.WriteLine($"Adding text block before match (length: {textContent.Length})");
                        blocks.Add(new TextBlock
                        {
                            Type = BlockType.Text,
                            Content = textContent
                        });
                    }
                }

                // Identify and add block type
                var matchContent = match.Value.Trim();
                if (match.Value.StartsWith("<table", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Adding table block");
                    blocks.Add(new TextBlock { Type = BlockType.Table, Content = matchContent });
                }
                else if (match.Value.StartsWith("<ul", StringComparison.OrdinalIgnoreCase) || match.Value.StartsWith("<ol", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Adding list block");
                    blocks.Add(new TextBlock { Type = BlockType.List, Content = matchContent });
                }

                // Update position to end of current match
                currentPosition = match.Index + match.Length;
                Console.WriteLine($"Updated current position to: {currentPosition}");
            }

            // Add remaining text after last match if any
            if (currentPosition < inputText.Length)
            {
                var remainingText = inputText.Substring(currentPosition).Trim();
                if (!string.IsNullOrWhiteSpace(remainingText))
                {
                    Console.WriteLine($"Adding remaining text block (length: {remainingText.Length})");
                    blocks.Add(new TextBlock { Type = BlockType.Text, Content = remainingText });
                }
            }

            Console.WriteLine($"Text segmentation complete. Created {blocks.Count} blocks:");
            for (int i = 0; i < blocks.Count; i++)
            {
                Console.WriteLine($"Block {i + 1}:");
                Console.WriteLine($"Type: {blocks[i].Type}");
                Console.WriteLine($"Content Length: {blocks[i].Content.Length}");
                Console.WriteLine($"Content:\n{blocks[i].Content}\n");
            }

            return blocks;
        }
    }
}