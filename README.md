# Document Processor

A C# command line tool for processing Word documents with Azure DevOps integration and dynamic content generation.

## Features

- Process Word documents with special tags
- Integrate with Azure DevOps work items
- Execute stored queries and format results
- Generate acronym tables automatically
- Convert HTML content to Word format

## Requirements

- .NET 8.0
- Azure DevOps connection settings configured in appsettings.json

## Setup

1. Configure Azure DevOps settings:
   ```bash
   # Copy the template configuration file
   cp appsettings.template.json appsettings.json
   ```

   Then edit appsettings.json with your Azure DevOps settings:
   ```json
   {
     "AzureDevOps": {
       "Organization": "your-organization-name",
       "PersonalAccessToken": "your-personal-access-token",
       "BaseUrl": "https://dev.azure.com"
     }
   }
   ```

   The application uses .NET's built-in configuration system (ConfigurationManager) to handle settings,
   making it easy to extend with additional configuration sources if needed.

2. Build the project:
   ```bash
   dotnet restore
   dotnet build
   ```

## Usage

### Create a test document
```bash
dotnet run -- --create-test
```

### Process a document
```bash
dotnet run -- <source_file> <output_file>
```

## Supported Tags

- `[[WorkItem:xxxx]]` - Retrieves and inserts content from a work item
- `[[QueryID:xxxx]]` - Executes a stored query and formats results
- `[[AcronymTable:true]]` - Generates a table of detected acronyms

## Development

Run tests:
```bash
dotnet test
```

## Configuration

The application uses appsettings.json for all configuration:

- AzureDevOps.Organization: Your Azure DevOps organization name
- AzureDevOps.PersonalAccessToken: Your Azure DevOps Personal Access Token
- AzureDevOps.BaseUrl: Base URL for Azure DevOps API (default: https://dev.azure.com)

Configuration is managed through Microsoft.Extensions.Configuration, supporting:
- JSON configuration files (appsettings.json)
- Environment variables
- Command-line arguments

Note: Make sure to keep your appsettings.json file secure and never commit it to version control.