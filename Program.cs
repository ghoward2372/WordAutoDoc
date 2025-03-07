using DocumentProcessor.Models;
using DocumentProcessor.Models.Configuration;
using DocumentProcessor.Services;
using DocumentProcessor.Tests;
using Microsoft.Extensions.Configuration;
using Microsoft.VisualStudio.Services.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;

namespace DocumentProcessor
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {

                Console.WriteLine("Starting Document Processor : " + args.Length + " arguments passed.");

                args.ForEach<string>(arg => Console.WriteLine("Argument : " + arg));

                // Check for help option
                if (args.Length >= 1 && args[0].Equals("/help", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Usage:");
                    Console.WriteLine("  DocumentProcessor.exe <source_file> <output_file> [/check]      - Process a document with optional grammar checking");
                    Console.WriteLine("  DocumentProcessor.exe --create-test                               - Create a test document");
                    Console.WriteLine("  DocumentProcessor.exe <path_to_scan> <output_file> /sbom          - Run SBOM Generation");
                    return;
                }

                // Check for create-test mode
                if (args.Length == 1 && args[0] == "--create-test")
                {
                    string testFilePath = "test_document.docx";
                    TestDocumentGenerator.CreateTestDocument(testFilePath);
                    Console.WriteLine($"Test document created at: {testFilePath}");
                    return;
                }

                // Ensure we have at least 2 arguments for document processing or SBOM generation
                if (args.Length < 2)
                {
                    Console.WriteLine("Insufficient arguments provided.");
                    Console.WriteLine("Usage:");
                    Console.WriteLine("  DocumentProcessor.exe <source_file> <output_file> [/check]");
                    Console.WriteLine("  DocumentProcessor.exe --create-test");
                    Console.WriteLine("  DocumentProcessor.exe <path_to_scan> <output_file> /sbom");
                    return;
                }

                // Check for SBOM Generation mode (third argument is "/sbom")
                if (args.Length > 2 && args[2].Equals("/sbom", StringComparison.OrdinalIgnoreCase))
                {
                    string pathToScan = args[0];
                    string outputPath = args[1];
                    Console.WriteLine("\n=== SBOM Generation Mode ===");
                    Console.WriteLine($"Path to scan: {pathToScan}");
                    Console.WriteLine($"Output path: {outputPath}");
                    var configuration = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                        .Build();
                    SBOMGenerator _sbomGenerator = new SBOMGenerator(pathToScan, configuration);
                    string sbomJson = await _sbomGenerator.GenerateSBOMAsync();
                    File.WriteAllText(outputPath, sbomJson);
                    Console.WriteLine("SBOM File saved to : " + outputPath);
                    Console.WriteLine("Have a nice day!");
                    return;
                }

                try
                {
                    Assembly interopWord = Assembly.Load("Microsoft.Office.Interop.Word");
                    Console.WriteLine("Interop loaded successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Interop failed to load: {ex.Message}");
                }
                // Otherwise, assume Document Processing mode.
                string sourceFile = args[0];
                string outputFile = args[1];
                bool checkGrammar = args.Length > 2 && args[2].Equals("/check", StringComparison.OrdinalIgnoreCase);

                Console.WriteLine("\n=== Document Processing Started ===");
                Console.WriteLine($"Source: {sourceFile}");
                Console.WriteLine($"Output: {outputFile}");
                Console.WriteLine($"Grammar check: {(checkGrammar ? "Enabled" : "Disabled")}");



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

                ReferenceDocProcessor refDocProcess = new ReferenceDocProcessor(adoService);
                refDocProcess.Intialize();

                var options = new DocumentProcessingOptions
                {
                    SourcePath = sourceFile,
                    OutputPath = outputFile,
                    AzureDevOpsService = adoService,
                    AcronymProcessor = new AcronymProcessor(acronymConfig),
                    HtmlConverter = new HtmlToWordConverter(),
                    ReferenceDocProcessor = refDocProcess,
                    FQDocumentField = config.FQDocumentFieldName
                };

                var processor = new WordDocumentProcessor(options);
                await processor.ProcessDocumentAsync();

                // Run grammar check if requested
                if (checkGrammar)
                {
                    try
                    {
                        var grammarChecker = new GrammarChecker(outputFile);
                        grammarChecker.CheckAndFixGrammar();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"\nWarning: Grammar checking failed - {ex.Message}");
                        Console.WriteLine("Continuing with document processing...");
                    }
                }

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