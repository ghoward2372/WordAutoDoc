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

            // Find all table matches in the text
            var tableMatches = Regex.Matches(
                inputText, 
                @"<table[^>]*>.*?</table>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase
            );

            Console.WriteLine($"Found {tableMatches.Count} table(s) in text");

            foreach (Match tableMatch in tableMatches)
            {
                Console.WriteLine($"Processing table match at position {tableMatch.Index}, length {tableMatch.Length}");

                // Add text before the table if any
                if (tableMatch.Index > currentPosition)
                {
                    var textContent = inputText.Substring(currentPosition, tableMatch.Index - currentPosition).Trim();
                    if (!string.IsNullOrWhiteSpace(textContent))
                    {
                        Console.WriteLine($"Adding text block before table (length: {textContent.Length})");
                        blocks.Add(new TextBlock
                        {
                            Type = BlockType.Text,
                            Content = textContent
                        });
                    }
                }

                // Add the table block
                var tableContent = tableMatch.Value.Trim();
                Console.WriteLine($"Adding table block (length: {tableContent.Length})");
                blocks.Add(new TextBlock
                {
                    Type = BlockType.Table,
                    Content = tableContent
                });

                // Update position to end of current table
                currentPosition = tableMatch.Index + tableMatch.Length;
                Console.WriteLine($"Updated current position to: {currentPosition}");
            }

            // Add remaining text after last table if any
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
                Console.WriteLine($"Content:\n{blocks[i].Content}\n");
            }

            return blocks;
        }
    }
}