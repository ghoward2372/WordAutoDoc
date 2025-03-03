using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace DocumentProcessor.Services
{
    public class TextBlockProcessor : ITextBlockProcessor
    {
        public enum BlockType
        {
            Text,
            Table,
            List
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

            // Find all special content (tables and lists) in the text
            var tableMatches = Regex.Matches(
                inputText, 
                @"<table[^>]*>.*?</table>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase
            );

            var listMatches = Regex.Matches(
                inputText,
                @"<ul[^>]*>.*?</ul>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase
            );

            Console.WriteLine($"Found {tableMatches.Count} table(s) and {listMatches.Count} list(s) in text");

            // Combine and sort all matches by their position in the text
            var allMatches = new List<(int Index, int Length, BlockType Type, string Content)>();

            foreach (Match match in tableMatches)
            {
                Console.WriteLine($"Table match at position {match.Index}, length {match.Length}");
                allMatches.Add((match.Index, match.Length, BlockType.Table, match.Value));
            }

            foreach (Match match in listMatches)
            {
                Console.WriteLine($"List match at position {match.Index}, length {match.Length}");
                allMatches.Add((match.Index, match.Length, BlockType.List, match.Value));
            }

            allMatches.Sort((a, b) => a.Index.CompareTo(b.Index));

            // Process all blocks in order
            foreach (var match in allMatches)
            {
                // Add text before the special block if any
                if (match.Index > currentPosition)
                {
                    var textContent = inputText.Substring(currentPosition, match.Index - currentPosition).Trim();
                    if (!string.IsNullOrWhiteSpace(textContent))
                    {
                        Console.WriteLine($"Adding text block before special content (length: {textContent.Length})");
                        blocks.Add(new TextBlock
                        {
                            Type = BlockType.Text,
                            Content = textContent
                        });
                    }
                }

                // Add the special block
                Console.WriteLine($"Adding {match.Type} block (length: {match.Content.Length})");
                blocks.Add(new TextBlock
                {
                    Type = match.Type,
                    Content = match.Content.Trim()
                });

                // Update position to end of current special block
                currentPosition = match.Index + match.Length;
            }

            // Add remaining text after last special block if any
            if (currentPosition < inputText.Length)
            {
                var remainingText = inputText.Substring(currentPosition).Trim();
                if (!string.IsNullOrWhiteSpace(remainingText))
                {
                    Console.WriteLine($"Adding remaining text block (length: {remainingText.Length})");
                    blocks.Add(new TextBlock
                    {
                        Type = BlockType.Text,
                        Content = remainingText
                    });
                }
            }

            Console.WriteLine($"Text segmentation complete. Created {blocks.Count} blocks:");
            for (int i = 0; i < blocks.Count; i++)
            {
                Console.WriteLine($"Block {i + 1}:");
                Console.WriteLine($"Type: {blocks[i].Type}");
                Console.WriteLine($"Content Length: {blocks[i].Content.Length}");
                Console.WriteLine($"Content Preview: {blocks[i].Content.Substring(0, Math.Min(100, blocks[i].Content.Length))}...\n");
            }

            return blocks;
        }
    }
}