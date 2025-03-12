import os
import json
import logging
import asyncio
from datetime import datetime
from typing import Any, Sequence
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from mcp.server import Server
from mcp.types import (
    Tool,
    TextContent,
    LoggingLevel,
    EmptyResult,
)

# .envファイルから環境変数を読み込む
load_dotenv()

# ロギング設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("sheet_mcp")

# Google Sheets APIの認証情報ファイルパスを環境変数から取得
CREDENTIALS_FILE = os.getenv("GOOGLE_SHEETS_CREDENTIALS_FILE")

if not CREDENTIALS_FILE:
    raise ValueError("Google Sheets credentials file path is required")

# 認証スコープ
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

try:
    # 認証情報を設定
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPES)
    # gspreadクライアントの初期化
    client = gspread.authorize(credentials)
    logger.info("Successfully authenticated with Google Sheets API")
except Exception as e:
    logger.error(f"Error authenticating with Google Sheets API: {str(e)}")
    raise RuntimeError(f"Authentication error: {str(e)}")

server = Server("sheet_mcp")

# ツールのリストアップ
@server.list_tools()
async def list_tools() -> list[Tool]:
    """List available tools for interacting with Google Sheets."""
    return [
        Tool(
            name="list_spreadsheets",
            description="List all available spreadsheets",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": [],
            },
        ),
        Tool(
            name="open_spreadsheet",
            description="Open a spreadsheet by title or URL",
            inputSchema={
                "type": "object",
                "properties": {
                    "identifier": {
                        "type": "string",
                        "description": "The title or URL of the spreadsheet to open",
                    },
                },
                "required": ["identifier"],
            },
        ),
        Tool(
            name="list_worksheets",
            description="List all worksheets in the currently opened spreadsheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                },
                "required": ["spreadsheet_id"],
            },
        ),
        Tool(
            name="read_worksheet",
            description="Read data from a worksheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "worksheet_name": {
                        "type": "string",
                        "description": "The name of the worksheet to read",
                    },
                    "range": {
                        "type": "string",
                        "description": "The range to read (e.g., 'A1:D10'). Optional, defaults to entire worksheet.",
                    },
                },
                "required": ["spreadsheet_id", "worksheet_name"],
            },
        ),
        Tool(
            name="update_cell",
            description="Update a single cell in a worksheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "worksheet_name": {
                        "type": "string",
                        "description": "The name of the worksheet",
                    },
                    "cell": {
                        "type": "string",
                        "description": "The cell reference (e.g., 'A1')",
                    },
                    "value": {
                        "type": ["string", "number", "boolean"],
                        "description": "The value to write to the cell",
                    },
                },
                "required": ["spreadsheet_id", "worksheet_name", "cell", "value"],
            },
        ),
        Tool(
            name="update_range",
            description="Update a range of cells in a worksheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "worksheet_name": {
                        "type": "string",
                        "description": "The name of the worksheet",
                    },
                    "range": {
                        "type": "string",
                        "description": "The range to update (e.g., 'A1:B2')",
                    },
                    "values": {
                        "type": "array",
                        "items": {
                            "type": "array",
                            "items": {
                                "type": ["string", "number", "boolean", "null"],
                            },
                        },
                        "description": "The values to write to the range (2D array)",
                    },
                },
                "required": ["spreadsheet_id", "worksheet_name", "range", "values"],
            },
        ),
        Tool(
            name="append_row",
            description="Append a row to a worksheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "worksheet_name": {
                        "type": "string",
                        "description": "The name of the worksheet",
                    },
                    "values": {
                        "type": "array",
                        "items": {
                            "type": ["string", "number", "boolean", "null"],
                        },
                        "description": "The values to append as a new row",
                    },
                },
                "required": ["spreadsheet_id", "worksheet_name", "values"],
            },
        ),
        Tool(
            name="create_spreadsheet",
            description="Create a new spreadsheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "title": {
                        "type": "string",
                        "description": "The title for the new spreadsheet",
                    },
                },
                "required": ["title"],
            },
        ),
        Tool(
            name="create_worksheet",
            description="Create a new worksheet in an existing spreadsheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "title": {
                        "type": "string",
                        "description": "The title for the new worksheet",
                    },
                    "rows": {
                        "type": "integer",
                        "description": "Number of rows (default: 100)",
                    },
                    "cols": {
                        "type": "integer",
                        "description": "Number of columns (default: 26)",
                    },
                },
                "required": ["spreadsheet_id", "title"],
            },
        ),
    ]

@server.call_tool()
async def call_tool(name: str, arguments: Any) -> Sequence[TextContent]:
    """Handle tool calls for Google Sheets operations."""
    if name == "list_spreadsheets":
        return await handle_list_spreadsheets(arguments)
    elif name == "open_spreadsheet":
        return await handle_open_spreadsheet(arguments)
    elif name == "list_worksheets":
        return await handle_list_worksheets(arguments)
    elif name == "read_worksheet":
        return await handle_read_worksheet(arguments)
    elif name == "update_cell":
        return await handle_update_cell(arguments)
    elif name == "update_range":
        return await handle_update_range(arguments)
    elif name == "append_row":
        return await handle_append_row(arguments)
    elif name == "create_spreadsheet":
        return await handle_create_spreadsheet(arguments)
    elif name == "create_worksheet":
        return await handle_create_worksheet(arguments)
    else:
        raise ValueError(f"Unknown tool: {name}")

async def handle_list_spreadsheets(arguments: Any) -> Sequence[TextContent]:
    """
    利用可能なすべてのスプレッドシートを一覧表示する。
    """
    try:
        spreadsheet_list = client.list_spreadsheet_files()
        spreadsheets = []
        
        for spreadsheet in spreadsheet_list:
            spreadsheets.append({
                "id": spreadsheet['id'],
                "name": spreadsheet['name'],
                "url": f"https://docs.google.com/spreadsheets/d/{spreadsheet['id']}"
            })
        
        return [
            TextContent(
                type="text",
                text=json.dumps(spreadsheets, indent=2, ensure_ascii=False),
            )
        ]
    except Exception as e:
        logger.error(f"Error listing spreadsheets: {str(e)}")
        raise RuntimeError(f"Error listing spreadsheets: {str(e)}")

async def handle_open_spreadsheet(arguments: Any) -> Sequence[TextContent]:
    """
    タイトルまたはURLでスプレッドシートを開く。
    - 引数 arguments: {"identifier": "<スプレッドシートのタイトルまたはURL>"}
    """
    if not isinstance(arguments, dict) or "identifier" not in arguments:
        raise ValueError("Invalid arguments for open_spreadsheet")
    
    identifier = arguments["identifier"]
    
    try:
        # URLかどうかをチェック
        if "https://docs.google.com/spreadsheets/d/" in identifier:
            # URLからIDを抽出
            sheet_id = identifier.split("/d/")[1].split("/")[0]
            spreadsheet = client.open_by_key(sheet_id)
        else:
            # タイトルで検索
            spreadsheet = client.open(identifier)
        
        spreadsheet_info = {
            "id": spreadsheet.id,
            "title": spreadsheet.title,
            "url": f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}",
            "sheet_count": len(spreadsheet.worksheets())
        }
        
        return [
            TextContent(
                type="text",
                text=json.dumps(spreadsheet_info, indent=2, ensure_ascii=False),
            )
        ]
    except gspread.exceptions.SpreadsheetNotFound:
        return [
            TextContent(
                type="text",
                text=f"Spreadsheet '{identifier}' not found.",
            )
        ]
    except Exception as e:
        logger.error(f"Error opening spreadsheet: {str(e)}")
        raise RuntimeError(f"Error opening spreadsheet: {str(e)}")

async def handle_list_worksheets(arguments: Any) -> Sequence[TextContent]:
    """
    スプレッドシート内のすべてのワークシートを一覧表示する。
    - 引数 arguments: {"spreadsheet_id": "<スプレッドシートID>"}
    """
    if not isinstance(arguments, dict) or "spreadsheet_id" not in arguments:
        raise ValueError("Invalid arguments for list_worksheets")
    
    spreadsheet_id = arguments["spreadsheet_id"]
    
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheets = []
        
        for worksheet in spreadsheet.worksheets():
            worksheets.append({
                "id": worksheet.id,
                "title": worksheet.title,
                "rows": worksheet.row_count,
                "cols": worksheet.col_count
            })
        
        return [
            TextContent(
                type="text",
                text=json.dumps(worksheets, indent=2, ensure_ascii=False),
            )
        ]
    except gspread.exceptions.SpreadsheetNotFound:
        return [
            TextContent(
                type="text",
                text=f"Spreadsheet with ID '{spreadsheet_id}' not found.",
            )
        ]
    except Exception as e:
        logger.error(f"Error listing worksheets: {str(e)}")
        raise RuntimeError(f"Error listing worksheets: {str(e)}")

async def handle_read_worksheet(arguments: Any) -> Sequence[TextContent]:
    """
    ワークシートからデータを読み込む。
    - 引数 arguments: {
        "spreadsheet_id": "<スプレッドシートID>",
        "worksheet_name": "<ワークシート名>",
        "range": "<範囲>" (オプション)
    }
    """
    if not isinstance(arguments, dict) or "spreadsheet_id" not in arguments or "worksheet_name" not in arguments:
        raise ValueError("Invalid arguments for read_worksheet")
    
    spreadsheet_id = arguments["spreadsheet_id"]
    worksheet_name = arguments["worksheet_name"]
    cell_range = arguments.get("range", None)
    
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        if cell_range:
            values = worksheet.get(cell_range)
        else:
            values = worksheet.get_all_values()
        
        return [
            TextContent(
                type="text",
                text=json.dumps(values, indent=2, ensure_ascii=False),
            )
        ]
    except gspread.exceptions.SpreadsheetNotFound:
        return [
            TextContent(
                type="text",
                text=f"Spreadsheet with ID '{spreadsheet_id}' not found.",
            )
        ]
    except gspread.exceptions.WorksheetNotFound:
        return [
            TextContent(
                type="text",
                text=f"Worksheet '{worksheet_name}' not found in spreadsheet.",
            )
        ]
    except Exception as e:
        logger.error(f"Error reading worksheet: {str(e)}")
        raise RuntimeError(f"Error reading worksheet: {str(e)}")

async def handle_update_cell(arguments: Any) -> Sequence[TextContent]:
    """
    ワークシートの単一セルを更新する。
    - 引数 arguments: {
        "spreadsheet_id": "<スプレッドシートID>",
        "worksheet_name": "<ワークシート名>",
        "cell": "<セル参照（例：A1）>",
        "value": <更新する値>
    }
    """
    if not isinstance(arguments, dict) or not all(k in arguments for k in ["spreadsheet_id", "worksheet_name", "cell", "value"]):
        raise ValueError("Invalid arguments for update_cell")
    
    spreadsheet_id = arguments["spreadsheet_id"]
    worksheet_name = arguments["worksheet_name"]
    cell = arguments["cell"]
    value = arguments["value"]
    
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        worksheet.update(cell, value)
        
        return [
            TextContent(
                type="text",
                text=f"Successfully updated cell {cell} in worksheet '{worksheet_name}' to value: {value}",
            )
        ]
    except gspread.exceptions.SpreadsheetNotFound:
        return [
            TextContent(
                type="text",
                text=f"Spreadsheet with ID '{spreadsheet_id}' not found.",
            )
        ]
    except gspread.exceptions.WorksheetNotFound:
        return [
            TextContent(
                type="text",
                text=f"Worksheet '{worksheet_name}' not found in spreadsheet.",
            )
        ]
    except Exception as e:
        logger.error(f"Error updating cell: {str(e)}")
        raise RuntimeError(f"Error updating cell: {str(e)}")

async def handle_update_range(arguments: Any) -> Sequence[TextContent]:
    """
    ワークシートの範囲を更新する。
    - 引数 arguments: {
        "spreadsheet_id": "<スプレッドシートID>",
        "worksheet_name": "<ワークシート名>",
        "range": "<範囲（例：A1:B2）>",
        "values": [[値1, 値2], [値3, 値4]] // 2次元配列
    }
    """
    if not isinstance(arguments, dict) or not all(k in arguments for k in ["spreadsheet_id", "worksheet_name", "range", "values"]):
        raise ValueError("Invalid arguments for update_range")
    
    spreadsheet_id = arguments["spreadsheet_id"]
    worksheet_name = arguments["worksheet_name"]
    cell_range = arguments["range"]
    values = arguments["values"]
    
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        worksheet.update(cell_range, values)
        
        return [
            TextContent(
                type="text",
                text=f"Successfully updated range {cell_range} in worksheet '{worksheet_name}'",
            )
        ]
    except gspread.exceptions.SpreadsheetNotFound:
        return [
            TextContent(
                type="text",
                text=f"Spreadsheet with ID '{spreadsheet_id}' not found.",
            )
        ]
    except gspread.exceptions.WorksheetNotFound:
        return [
            TextContent(
                type="text",
                text=f"Worksheet '{worksheet_name}' not found in spreadsheet.",
            )
        ]
    except Exception as e:
        logger.error(f"Error updating range: {str(e)}")
        raise RuntimeError(f"Error updating range: {str(e)}")

async def handle_append_row(arguments: Any) -> Sequence[TextContent]:
    """
    ワークシートに行を追加する。
    - 引数 arguments: {
        "spreadsheet_id": "<スプレッドシートID>",
        "worksheet_name": "<ワークシート名>",
        "values": [値1, 値2, 値3, ...] // 1次元配列
    }
    """
    if not isinstance(arguments, dict) or not all(k in arguments for k in ["spreadsheet_id", "worksheet_name", "values"]):
        raise ValueError("Invalid arguments for append_row")
    
    spreadsheet_id = arguments["spreadsheet_id"]
    worksheet_name = arguments["worksheet_name"]
    values = arguments["values"]
    
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        result = worksheet.append_row(values)
        
        return [
            TextContent(
                type="text",
                text=f"Successfully appended row to worksheet '{worksheet_name}'",
            )
        ]
    except gspread.exceptions.SpreadsheetNotFound:
        return [
            TextContent(
                type="text",
                text=f"Spreadsheet with ID '{spreadsheet_id}' not found.",
            )
        ]
    except gspread.exceptions.WorksheetNotFound:
        return [
            TextContent(
                type="text",
                text=f"Worksheet '{worksheet_name}' not found in spreadsheet.",
            )
        ]
    except Exception as e:
        logger.error(f"Error appending row: {str(e)}")
        raise RuntimeError(f"Error appending row: {str(e)}")

async def handle_create_spreadsheet(arguments: Any) -> Sequence[TextContent]:
    """
    新しいスプレッドシートを作成する。
    - 引数 arguments: {"title": "<新しいスプレッドシートのタイトル>"}
    """
    if not isinstance(arguments, dict) or "title" not in arguments:
        raise ValueError("Invalid arguments for create_spreadsheet")
    
    title = arguments["title"]
    
    try:
        spreadsheet = client.create(title)
        
        spreadsheet_info = {
            "id": spreadsheet.id,
            "title": spreadsheet.title,
            "url": f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}"
        }
        
        return [
            TextContent(
                type="text",
                text=json.dumps(spreadsheet_info, indent=2, ensure_ascii=False),
            )
        ]
    except Exception as e:
        logger.error(f"Error creating spreadsheet: {str(e)}")
        raise RuntimeError(f"Error creating spreadsheet: {str(e)}")

async def handle_create_worksheet(arguments: Any) -> Sequence[TextContent]:
    """
    既存のスプレッドシートに新しいワークシートを作成する。
    - 引数 arguments: {
        "spreadsheet_id": "<スプレッドシートID>",
        "title": "<新しいワークシートのタイトル>",
        "rows": <行数> (オプション、デフォルト: 100),
        "cols": <列数> (オプション、デフォルト: 26)
    }
    """
    if not isinstance(arguments, dict) or not all(k in arguments for k in ["spreadsheet_id", "title"]):
        raise ValueError("Invalid arguments for create_worksheet")
    
    spreadsheet_id = arguments["spreadsheet_id"]
    title = arguments["title"]
    rows = arguments.get("rows", 100)
    cols = arguments.get("cols", 26)
    
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)
        
        worksheet_info = {
            "id": worksheet.id,
            "title": worksheet.title,
            "rows": worksheet.row_count,
            "cols": worksheet.col_count
        }
        
        return [
            TextContent(
                type="text",
                text=json.dumps(worksheet_info, indent=2, ensure_ascii=False),
            )
        ]
    except gspread.exceptions.SpreadsheetNotFound:
        return [
            TextContent(
                type="text",
                text=f"Spreadsheet with ID '{spreadsheet_id}' not found.",
            )
        ]
    except Exception as e:
        logger.error(f"Error creating worksheet: {str(e)}")
        raise RuntimeError(f"Error creating worksheet: {str(e)}")

async def main():
    import mcp
    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())