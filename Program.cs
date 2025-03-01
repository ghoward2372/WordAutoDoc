using DocumentProcessor.Models;
using DocumentProcessor.Services;
using DocumentProcessor.Tests;
using System;
using System.Threading.Tasks;

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
                    AcronymProcessor = new AcronymProcessor(),
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
    }
}