using DocumentProcessor.Models;
using DocumentProcessor.Services;
using DocumentProcessor.Tests;
using DocumentProcessor.Models.Configuration;
using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;
using System.IO;

namespace DocumentProcessor
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                if (args.Length == 1 && args[0] == "--create-test")
                {
                    string testFilePath = "test_document.docx";
                    TestDocumentGenerator.CreateTestDocument(testFilePath);
                    Console.WriteLine($"Test document created at: {testFilePath}");
                    return;
                }

                if (args.Length != 2)
                {
                    Console.WriteLine("Usage: DocumentProcessor.exe <source_file> <output_file>");
                    Console.WriteLine("       DocumentProcessor.exe --create-test");
                    return;
                }

                string sourceFile = args[0];
                string outputFile = args[1];

                Console.WriteLine($"\n=== Document Processing Started ===");
                Console.WriteLine($"Source: {sourceFile}");
                Console.WriteLine($"Output: {outputFile}");

                // Load acronym configuration
                var acronymConfig = LoadAcronymConfiguration();
                Console.WriteLine($"Loaded {acronymConfig.KnownAcronyms.Count} known acronyms and {acronymConfig.IgnoredAcronyms.Count} ignored acronyms");

                // Initialize services with configuration
                IAzureDevOpsService? adoService = null;
                try
                {
                    adoService = AzureDevOpsService.Initialize();
                    Console.WriteLine("Successfully connected to Azure DevOps.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Azure DevOps integration not available - {ex.Message}");
                    Console.WriteLine("Continuing with limited functionality (acronym processing only).");
                }
                var config = ConfigurationService.LoadAzureDevOpsConfig();

                var options = new DocumentProcessingOptions
                {
                    SourcePath = sourceFile,
                    OutputPath = outputFile,
                    AzureDevOpsService = adoService,
                    AcronymProcessor = new AcronymProcessor(acronymConfig),
                    HtmlConverter = new HtmlToWordConverter(),
                    FQDocumentField = config.FQDocumentFieldName
                };

                var processor = new WordDocumentProcessor(options);
                await processor.ProcessDocumentAsync();

                Console.WriteLine("\n=== Processing Complete ===");
                Console.WriteLine($"Output document ready at: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n=== Processing Failed ===");
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
                Environment.Exit(1);
            }
        }

        private static AcronymConfiguration LoadAcronymConfiguration()
        {
            try
            {
                string configPath = "acronyms.json";
                if (!File.Exists(configPath))
                {
                    Console.WriteLine("Warning: acronyms.json not found, using empty configuration");
                    return new AcronymConfiguration
                    {
                        KnownAcronyms = new Dictionary<string, string>(),
                        IgnoredAcronyms = new HashSet<string>()
                    };
                }

                string jsonContent = File.ReadAllText(configPath);
                var config = JsonSerializer.Deserialize<AcronymConfiguration>(jsonContent);
                if (config == null)
                {
                    throw new InvalidOperationException("Failed to deserialize acronym configuration");
                }
                return config;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading acronym configuration: {ex.Message}");
                return new AcronymConfiguration
                {
                    KnownAcronyms = new Dictionary<string, string>(),
                    IgnoredAcronyms = new HashSet<string>()
                };
            }
        }
    }
}