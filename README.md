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
- Azure DevOps connection settings configured in appsettings.json or environment variables

## Setup

1. Configure Azure DevOps settings using one of these methods:

   a. Using environment variables (recommended for security):
   ```bash
   # Set these environment variables
   export ADO_ORGANIZATION="your-organization-name"
   export ADO_PAT="your-personal-access-token"
   export ADO_BASEURL="https://dev.azure.com"  # Optional, defaults to https://dev.azure.com
   ```

   b. Using configuration file:
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

   Note: Environment variables take precedence over appsettings.json values.

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

The application supports multiple configuration sources in the following order of precedence:

1. Environment Variables (Most secure, recommended for production)
   - ADO_ORGANIZATION: Your Azure DevOps organization name
   - ADO_PAT: Your Azure DevOps Personal Access Token
   - ADO_BASEURL: Base URL for Azure DevOps API (optional)

2. Configuration File (appsettings.json)
   - AzureDevOps.Organization: Your Azure DevOps organization name
   - AzureDevOps.PersonalAccessToken: Your Azure DevOps Personal Access Token
   - AzureDevOps.BaseUrl: Base URL for Azure DevOps API (default: https://dev.azure.com)

Configuration is managed through Microsoft.Extensions.Configuration, supporting:
- Environment variables (with ADO_ prefix)
- JSON configuration files (appsettings.json)
- Command-line arguments

Security Note: 
- Always use environment variables for sensitive information in production environments
- Never commit appsettings.json containing real credentials to version control
- Ensure your Azure DevOps PAT has the minimum required permissions (read work items, execute queries)