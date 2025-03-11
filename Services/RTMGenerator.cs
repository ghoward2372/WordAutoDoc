using Microsoft.Extensions.Configuration;
using System;
using System.IO;

namespace DocumentProcessor.Services
{
    public class RTMGenerator
    {
        private IAzureDevOpsService _adoService;
        private IHtmlToWordConverter _htmlConverter;
        private string _fieldFQName;

        public RTMGenerator(IAzureDevOpsService adoService)
        {
            _adoService = adoService;
            _htmlConverter = new HtmlToWordConverter();

        }

        public void Intialize()
        {
            var configuration = new ConfigurationBuilder()
                  .SetBasePath(Directory.GetCurrentDirectory())
                  .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                  .Build();


            _fieldFQName = configuration["AzureDevOps:FQDocumentFieldName"] ?? throw new ArgumentNullException("FQ Document Field Name not found in configuration");


        }
    }
}
