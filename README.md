# Document Processor

A C# command line tool for processing Word documents with Azure DevOps integration and dynamic content generation.

## Features

- Process Word documents with special tags
- Integrate with Azure DevOps work items
- Execute stored queries and format results
- Generate acronym tables automatically
- Convert HTML content to Word format

## Requirements

- .NET 7.0
- Azure DevOps connection settings configured in appsettings.json

## Setup

1. Configure appsettings.json:
   ```json
   {
     "AzureDevOps": {
       "Organization": "your-organization-name",
       "PersonalAccessToken": "your-personal-access-token",
       "BaseUrl": "https://dev.azure.com"
     }
   }
   ```

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