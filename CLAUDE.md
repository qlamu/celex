# Celex - XLSX Library for Elixir

## Project Overview

Celex is a robust, stream-based Elixir library for reading and writing XLSX files. It's designed to handle large files (millions of rows) with constant memory usage.

## Quick Start

```elixir
# Builder API (small files)
Celex.new()
|> Celex.add_sheet("Sales")
|> Celex.put_cell("A1", "Product")
|> Celex.put_cell("B1", 42)
|> Celex.write!("report.xlsx")

# Streaming Write (large files)
Celex.stream_write("large.xlsx", fn writer ->
  writer
  |> Celex.StreamWriter.start_sheet("Data")
  |> Celex.StreamWriter.write_rows(data_stream)
end)

# Streaming Read
"large.xlsx"
|> Celex.stream_sheet("Data")
|> Stream.drop(1)  # Skip header
|> Enum.each(&process_row/1)
```

## Project Structure

```
lib/
├── celex.ex                      # Main public API
└── celex/
    ├── cell.ex                   # Cell struct (value, type, style, formula)
    ├── row.ex                    # Row struct for streaming
    ├── sheet.ex                  # Sheet struct with cells map
    ├── workbook.ex               # Workbook struct with sheets
    ├── style.ex                  # Style structs (Font, Fill, Border, Alignment)
    ├── stream_writer.ex          # Streaming XLSX writer (constant memory)
    ├── stream_reader.ex          # Streaming XLSX reader (constant memory)
    ├── reader.ex                 # Non-streaming reader (loads full workbook)
    ├── writer.ex                 # Non-streaming writer (uses StreamWriter)
    ├── xml/
    │   ├── content_types.ex      # [Content_Types].xml generation
    │   ├── relationships.ex      # .rels files generation
    │   ├── workbook_xml.ex       # xl/workbook.xml generation
    │   ├── worksheet_xml.ex      # xl/worksheets/sheetN.xml generation
    │   ├── shared_strings.ex     # xl/sharedStrings.xml generation
    │   └── styles_xml.ex         # xl/styles.xml generation
    └── utils/
        ├── cell_reference.ex     # A1 notation parsing (A1 -> {1, 1})
        └── date_utils.ex         # Excel serial date conversion
```

## Key Design Decisions

### Streaming Architecture
- **Writing**: Rows written directly to temp files, ZIP built at close()
- **Reading**: ZIP opened in memory mode, worksheets parsed with regex-based extraction
- **Memory**: O(1) constant - only current row in memory during streaming
- **Inline Strings**: Uses `<is><t>text</t></is>` instead of shared strings table for streaming mode

### XLSX File Format
XLSX files are ZIP archives containing:
- `[Content_Types].xml` - MIME types
- `_rels/.rels` - Root relationships
- `xl/workbook.xml` - Workbook definition
- `xl/worksheets/sheet1.xml` - Worksheet data
- `xl/styles.xml` - Cell styles
- `xl/sharedStrings.xml` - Deduplicated strings (optional)

### Cell Types
- `:string` - Text values (inline strings in XML)
- `:number` - Integer or float
- `:boolean` - true/false (stored as 1/0)
- `:date` - Elixir Date (converted to Excel serial number)
- `:datetime` - NaiveDateTime (serial number with fractional time)
- `:formula` - Formula string (without leading '=')
- `:blank` - nil value

## Development Commands

```bash
# Run all tests
mix test

# Run specific test file
mix test test/celex/cell_test.exs

# Run tests with coverage
mix test --cover

# Format code
mix format

# Generate docs
mix docs
```

## Testing Conventions

- All modules have corresponding test files in `test/celex/`
- Tests follow TDD pattern (tests written first)
- Use `async: true` for test modules when possible
- Temp files in tests use `System.tmp_dir!()` and cleanup in `on_exit`

## Dependencies

**Zero runtime dependencies** - Uses only Erlang's built-in modules:
- `:zip` - ZIP archive creation/extraction
- `:xmerl` - XML parsing (for future reader)

Dev dependencies:
- `ex_doc` - Documentation generation

## Common Patterns

### Creating cells with options
```elixir
Cell.new("Hello")                           # Simple value
Cell.new(42, style: %{bold: true})          # With style
Cell.new(nil, formula: "SUM(A1:A10)")       # Formula
```

### Working with cell references
```elixir
CellReference.parse("AA100")    # => {:ok, {100, 27}}
CellReference.to_string({1, 27}) # => "AA1"
CellReference.column_to_index("AA") # => 27
```

### Date conversion
```elixir
DateUtils.date_to_serial(~D[2000-01-01])  # => 36526
DateUtils.serial_to_date(36526)            # => ~D[2000-01-01]
```

### Reading files
```elixir
# Non-streaming (small files)
workbook = Celex.read!("data.xlsx")
{:ok, sheet} = Celex.get_sheet(workbook, "Sales")
rows = Celex.get_rows(sheet)
value = Celex.get_cell(sheet, "A1")

# Streaming (large files)
"large.xlsx"
|> Celex.stream_sheet("Data")
|> Stream.drop(1)  # Skip header
|> Enum.each(&process_row/1)
```

## Future Work

- [x] Streaming Reader (regex-based XML parsing)
- [x] Non-streaming Reader API (load full workbook)
- [x] Border styles (thin, medium, thick, dashed, dotted, double)
- [x] Number formats (built-in and custom)
- [x] Memory benchmarks for large files
- [ ] Image support (xl/drawings/, xl/media/)

## Benchmark Results

Run benchmarks with: `mix test --include benchmark`

| Operation | Rows | Duration | Memory Growth | Throughput |
|-----------|------|----------|---------------|------------|
| Write | 10,000 | 76ms | 0.01MB | ~130k rows/s |
| Write | 100,000 | 788ms | 15MB | ~127k rows/s |
| Write | 500,000 | 3.9s | 71MB | ~127k rows/s |
| Read | 10,000 | 199ms | 0.22MB | ~50k rows/s |
| Read | 100,000 | 1.95s | ~0MB | ~51k rows/s |
| Read | 500,000 | 9.8s | ~0MB | ~51k rows/s |

File size scales linearly at ~23-25 bytes per row.

**Note:** Reading uses true SAX-based streaming with O(1) memory usage.

## Excel Date System Note

Excel has a bug: it incorrectly treats 1900 as a leap year. February 29, 1900 (serial 60) doesn't exist but Excel thinks it does. The `DateUtils` module accounts for this to ensure compatibility.
