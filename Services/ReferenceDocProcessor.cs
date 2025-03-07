using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;


namespace DocumentProcessor.Services
{
    public class ReferenceDocProcessor
    {
        private Dictionary<string, string> _referenceDocs = new();
        private Dictionary<string, string> _knownReferenceDocs = new();
        private string _referenceDocRegex;

        private IAzureDevOpsService _azureService;


        public ReferenceDocProcessor(IAzureDevOpsService azureDevOpsService)
        {
            _azureService = azureDevOpsService;

        }
        public static string RemoveHtmlTags(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            // This regex matches anything between '<' and '>', non-greedily.
            return Regex.Replace(input, "<.*?>", string.Empty).Trim();
        }

        public async void Intialize()
        {
            var configuration = new ConfigurationBuilder()
                  .SetBasePath(Directory.GetCurrentDirectory())
                  .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                  .Build();

            string referenceWorkItemID = configuration["AzureDevOps:ReferenceSourceWorkItem"] ?? throw new ArgumentNullException("NIST API Key not found in configuration");
            _referenceDocRegex = configuration["AzureDevOps:ReferenceDocRegex"] ?? throw new ArgumentNullException("Reference Doc Regex not found in configuration");
            string fieldFQName = configuration["AzureDevOps:FQDocumentFieldName"] ?? throw new ArgumentNullException("FQ Document Field Name not found in configuration");

            if (string.IsNullOrEmpty(referenceWorkItemID) == false)
            {
                List<int> refIdList = new List<int>();
                refIdList.Add(Int32.Parse(referenceWorkItemID));

                var workItems = await _azureService.GetWorkItemsAsync(refIdList);

                if (workItems.Any())
                {
                    var wItem = workItems.First();
                    string referenceDocString = wItem.Fields[fieldFQName].ToString() ?? "";
                    if (string.IsNullOrEmpty(referenceDocString) == false)
                    {
                        string[] refDocs = referenceDocString.Split(',');
                        foreach (string refDoc in refDocs)
                        {
                            string[] refDocParts = refDoc.Split(':');
                            if (refDocParts.Length == 2)
                            {
                                _knownReferenceDocs.Add(RemoveHtmlTags(refDocParts[0]), RemoveHtmlTags(refDocParts[1]));
                            }
                        }
                    }
                }
            }
        }

        public void ProcessText(string text)
        {
            var referencePattern = new Regex(@"S3I-CAFRS-([^-]+)-([^-]+)-(\d+)", RegexOptions.Compiled);

            var matches = referencePattern.Matches(text);
            foreach (Match match in matches)
            {
                string reference = match.Groups[0].Value;
                if (!_referenceDocs.ContainsKey(reference))
                {
                    if (_knownReferenceDocs.TryGetValue(reference, out var knownDefinition))
                    {
                        _referenceDocs[reference] = knownDefinition;
                    }
                    else
                    {
                        _referenceDocs[reference] = string.Empty; // Store empty definition
                    }
                }
            }
        }

        public string[][] GetReferenceTableData()
        {
            // Create header row
            var tableData = new List<string[]>
            {
                new[] { "Document Number", "Document Title" }
            };

            // Add acronym rows, sorted alphabetically
            tableData.AddRange(
                _referenceDocs
                    .OrderBy(a => a.Key)
                    .Select(a => new[] { a.Key, a.Value })
            );

            return tableData.ToArray();
        }
    }
}
