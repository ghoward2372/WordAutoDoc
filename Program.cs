using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using DocumentProcessor.Services;
using DocumentProcessor.Models;
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

                // Initialize Azure DevOps connection
                var organization = Environment.GetEnvironmentVariable("ADO_ORGANIZATION");
                var pat = Environment.GetEnvironmentVariable("ADO_PAT");

                if (string.IsNullOrEmpty(organization) || string.IsNullOrEmpty(pat))
                {
                    throw new Exception("Azure DevOps organization and PAT must be set in environment variables.");
                }

                var credentials = new VssBasicCredential(string.Empty, pat);
                var connection = new VssConnection(new Uri($"https://dev.azure.com/{organization}"), credentials);
                var witClient = connection.GetClient<WorkItemTrackingHttpClient>();

                // Initialize services
                var azureDevOpsService = new AzureDevOpsService(witClient);
                var acronymProcessor = new AcronymProcessor();
                var htmlConverter = new HtmlToWordConverter();

                var options = new DocumentProcessingOptions
                {
                    SourcePath = sourceFile,
                    OutputPath = outputFile,
                    AzureDevOpsService = azureDevOpsService,
                    AcronymProcessor = acronymProcessor,
                    HtmlConverter = htmlConverter
                };

                var processor = new WordDocumentProcessor(options);
                await processor.ProcessDocumentAsync();

                Console.WriteLine("Document processing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
                Environment.Exit(1);
            }
        }
    }
}