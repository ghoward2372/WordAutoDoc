using DocumentProcessor.Services;
using Microsoft.Extensions.Configuration;
using System.IO;
using System.Threading.Tasks;

namespace DocumentProcessor.Models.TagProcessors
{
    class SBOMTagProcessor : ITagProcessor
    {
        private SBOMGenerator _sbomGenerator;

        public SBOMTagProcessor(IAzureDevOpsService azureDevOpsService, IHtmlToWordConverter htmlConverter)
        {
            AzureDevOpsService = azureDevOpsService;
            HtmlConverter = htmlConverter;
        }

        public IAzureDevOpsService AzureDevOpsService { get; }
        public IHtmlToWordConverter HtmlConverter { get; }

        public Task<ProcessingResult> ProcessTagAsync(string tagContent)
        {
            return ProcessTagAsync(tagContent, null);
        }

        public async Task<ProcessingResult> ProcessTagAsync(string tagContent, DocumentProcessingOptions? options)
        {
            var configuration = new ConfigurationBuilder()
                   .SetBasePath(Directory.GetCurrentDirectory())
                   .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                   .Build();
            _sbomGenerator = new SBOMGenerator(tagContent, configuration);
            string sbomJson = await _sbomGenerator.GenerateSBOMAsync();
            File.WriteAllText("C:\\temp\\sbom.json", sbomJson);
            return ProcessingResult.FromText(sbomJson);


        }
    }
}
