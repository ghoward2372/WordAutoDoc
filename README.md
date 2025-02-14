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
- Azure DevOps Organization and Personal Access Token (PAT)

## Setup

1. Set environment variables:
   ```bash
   ADO_ORGANIZATION=your-organization
   ADO_PAT=your-personal-access-token
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
