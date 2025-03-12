# Google Sheets MCP Server

[![smithery badge](https://smithery.ai/badge/sheets-mcp)](https://smithery.ai/server/sheets-mcp)

An MCP server to read, write, and manage Google Sheets directly through Claude chat.

## Features

- List all available spreadsheets
- Open spreadsheets by title or URL
- Create new spreadsheets and worksheets
- Read data from worksheets (entire sheet or specific ranges)
- Update individual cells or ranges of cells
- Append rows to worksheets

## Quick Setup

### Prerequisites

1. **Google Sheets API Credentials**:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project
   - Enable the Google Sheets API and Google Drive API
   - Create a Service Account
   - Download the JSON credentials file
   - Share your target Google Sheets with the service account email

### Installing via Smithery (Recommended)

To install Google Sheets MCP Server for Claude Desktop automatically via [Smithery](https://smithery.ai/server/sheets-mcp):

```bash
npx -y @smithery/cli install sheets-mcp --client claude
```

### Manual Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/sheets-mcp.git
```

2. Install UV globally using Homebrew in Terminal:
```bash
brew install uv
```

3. Create claude_desktop_config.json:
   - For MacOS: Open directory `~/Library/Application Support/Claude/` and create the file inside it
   - For Windows: Open directory `%APPDATA%/Claude/` and create the file inside it

4. Add this configuration to claude_desktop_config.json:
```json
{
  "mcpServers": {
    "sheets_mcp": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/sheets-mcp",
        "run",
        "sheets-mcp"
      ],
      "env": {
        "GOOGLE_SHEETS_CREDENTIALS_FILE": "/path/to/your-credentials.json"
      }
    }
  }
}
```

5. Update the config file:
   - Replace `/path/to/sheets-mcp` with your actual repository path
   - Replace `/path/to/your-credentials.json` with the path to your Google Service Account credentials JSON file

6. Quit Claude completely and reopen it

## Usage Examples

* "Show me all my available spreadsheets"
* "Open my Budget 2025 spreadsheet"
* "List all worksheets in my spreadsheet"
* "Read data from my Expenses worksheet"
* "Update cell A1 to 'Total Expenses' in my Budget worksheet"
* "Add a new row with today's expenses"
* "Create a new spreadsheet called 'Project Tracker'"

## Troubleshooting

If not working:
- Make sure UV is installed globally (if not, uninstall with `pip uninstall uv` and reinstall with `brew install uv`)
- Or find UV path with `which uv` and replace `"command": "uv"` with the full path
- Verify your Google Sheets API credentials file exists and is correctly referenced
- Check if the sheets-mcp path in config matches your actual repository location
- Ensure your Google Service Account has access to the spreadsheets you're trying to read/write

## API Reference

### Available Tools

1. **list_spreadsheets**
   - Lists all available spreadsheets
   - No parameters required

2. **open_spreadsheet**
   - Opens a spreadsheet by title or URL
   - Parameters: `identifier` (string) - The title or URL of the spreadsheet

3. **list_worksheets**
   - Lists all worksheets in a spreadsheet
   - Parameters: `spreadsheet_id` (string) - The ID of the spreadsheet

4. **read_worksheet**
   - Reads data from a worksheet
   - Parameters: 
     - `spreadsheet_id` (string) - The ID of the spreadsheet
     - `worksheet_name` (string) - The name of the worksheet
     - `range` (string, optional) - The range to read (e.g., 'A1:D10')

5. **update_cell**
   - Updates a single cell in a worksheet
   - Parameters:
     - `spreadsheet_id` (string) - The ID of the spreadsheet
     - `worksheet_name` (string) - The name of the worksheet
     - `cell` (string) - The cell reference (e.g., 'A1')
     - `value` (string/number/boolean) - The value to write

6. **update_range**
   - Updates a range of cells in a worksheet
   - Parameters:
     - `spreadsheet_id` (string) - The ID of the spreadsheet
     - `worksheet_name` (string) - The name of the worksheet
     - `range` (string) - The range to update (e.g., 'A1:B2')
     - `values` (2D array) - The values to write

7. **append_row**
   - Appends a row to a worksheet
   - Parameters:
     - `spreadsheet_id` (string) - The ID of the spreadsheet
     - `worksheet_name` (string) - The name of the worksheet
     - `values` (array) - The values to append as a new row

8. **create_spreadsheet**
   - Creates a new spreadsheet
   - Parameters: `title` (string) - The title for the new spreadsheet

9. **create_worksheet**
   - Creates a new worksheet in an existing spreadsheet
   - Parameters:
     - `spreadsheet_id` (string) - The ID of the spreadsheet
     - `title` (string) - The title for the new worksheet
     - `rows` (integer, optional) - Number of rows (default: 100)
     - `cols` (integer, optional) - Number of columns (default: 26)