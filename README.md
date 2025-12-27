# Celex

A robust, stream-based Elixir library for reading and writing XLSX files.

Designed to handle large files (millions of rows) with constant memory usage, while also providing a convenient builder API for smaller files.

## Features

- **Stream-based I/O** - Read and write files with millions of rows using constant O(1) memory
- **Builder pattern API** - Convenient chainable API for creating workbooks
- **Full XLSX support** - Styles, formulas, merged cells, column widths, row heights
- **Zero runtime dependencies** - Uses only Erlang's built-in `:zip` module
- **Comprehensive styling** - Fonts, fills, borders, alignment, number formats

## Installation

Add `celex` to your list of dependencies in `mix.exs`:

```elixir
def deps do
  [
    {:celex, "~> 0.1.0"}
  ]
end
```

## Quick Start

### Writing Files

#### Builder API (small to medium files)

```elixir
Celex.new()
|> Celex.add_sheet("Sales")
|> Celex.put_cell("A1", "Product", style: %{font: %{bold: true}})
|> Celex.put_cell("B1", "Revenue", style: %{font: %{bold: true}})
|> Celex.put_cell("A2", "Widget")
|> Celex.put_cell("B2", 1299.99)
|> Celex.put_cell("A3", "Gadget")
|> Celex.put_cell("B3", 849.50)
|> Celex.put_cell("B4", nil, formula: "SUM(B2:B3)")
|> Celex.write!("sales_report.xlsx")
```

#### Streaming API (large files)

```elixir
# Write millions of rows with constant memory
Celex.stream_write("large_export.xlsx", fn writer ->
  writer
  |> Celex.StreamWriter.start_sheet("Data")
  |> Celex.StreamWriter.write_row(["ID", "Name", "Value"])
  |> Celex.StreamWriter.write_rows(
    Stream.map(1..1_000_000, fn i ->
      [i, "Item #{i}", :rand.uniform(1000)]
    end)
  )
end)
```

### Reading Files

#### Non-streaming (small to medium files)

```elixir
workbook = Celex.read!("data.xlsx")

# Get sheet names
Celex.sheet_names(workbook)
# => ["Sheet1", "Sheet2"]

# Access a sheet
{:ok, sheet} = Celex.get_sheet(workbook, "Sheet1")

# Get all rows
rows = Celex.get_rows(sheet)
# => [["Name", "Age"], ["Alice", 30], ["Bob", 25]]

# Get individual cells
Celex.get_cell(sheet, "A1")  # => "Name"
Celex.get_cell(sheet, {2, 1}) # => "Alice" (row 2, column 1)
```

#### Streaming (large files)

```elixir
# Process rows one at a time with constant memory
"large_file.xlsx"
|> Celex.stream_sheet("Data")
|> Stream.drop(1)  # Skip header row
|> Stream.filter(fn row -> Enum.at(row, 2) > 100 end)
|> Enum.each(&process_row/1)
```

## Styling

Celex supports comprehensive cell styling:

```elixir
# Font styling
style = %{
  font: %{
    bold: true,
    italic: true,
    size: 14,
    color: "#FF0000",
    name: "Arial"
  }
}

# Background fill
style = %{
  fill: %{
    color: "#FFFF00",
    pattern: :solid
  }
}

# Borders
style = %{
  border: %{
    left: :thin,
    right: :thin,
    top: :medium,
    bottom: :medium,
    color: "#000000"
  }
}
# Border styles: :thin, :medium, :thick, :dashed, :dotted, :double

# Alignment
style = %{
  alignment: %{
    horizontal: :center,  # :left, :center, :right, :fill, :justify
    vertical: :middle,    # :top, :middle, :bottom
    wrap_text: true
  }
}

# Number formats
style = %{number_format: "#,##0.00"}
style = %{number_format: "0.00%"}
style = %{number_format: "yyyy-mm-dd"}

# Combined styling
Celex.put_cell(workbook, "A1", "Header", style: %{
  font: %{bold: true, size: 14},
  fill: %{color: "#4472C4", pattern: :solid},
  alignment: %{horizontal: :center}
})
```

## Cell Types

Celex automatically infers cell types, or you can specify them explicitly:

```elixir
# Automatic type inference
Celex.put_cell(wb, "A1", "Hello")      # String
Celex.put_cell(wb, "A2", 42)           # Integer
Celex.put_cell(wb, "A3", 3.14)         # Float
Celex.put_cell(wb, "A4", true)         # Boolean
Celex.put_cell(wb, "A5", ~D[2024-01-15]) # Date
Celex.put_cell(wb, "A6", ~N[2024-01-15 10:30:00]) # DateTime

# Formulas
Celex.put_cell(wb, "A7", nil, formula: "SUM(A2:A6)")
Celex.put_cell(wb, "A8", nil, formula: "IF(A4,\"Yes\",\"No\")")
```

## Advanced Features

### Merged Cells

```elixir
Celex.new()
|> Celex.add_sheet("Report")
|> Celex.put_cell("A1", "Quarterly Report")
|> Celex.merge_cells("A1:D1")
|> Celex.write!("report.xlsx")
```

### Column Widths and Row Heights

```elixir
workbook
|> Celex.set_column_width("A", 25)
|> Celex.set_column_width("B", 15)
|> Celex.set_row_height(1, 30)
```

### Multiple Sheets

```elixir
Celex.new()
|> Celex.add_sheet("Summary")
|> Celex.put_cell("A1", "Total")
|> Celex.put_cell("B1", nil, formula: "SUM(Data!A:A)")
|> Celex.add_sheet("Data")
|> Celex.put_cell("A1", 100)
|> Celex.put_cell("A2", 200)
|> Celex.write!("multi_sheet.xlsx")
```

### Low-level Streaming Control

```elixir
{:ok, reader} = Celex.stream_open("file.xlsx")

# Get available sheets
sheets = Celex.StreamReader.sheet_names(reader)

# Stream from specific sheet
Celex.StreamReader.stream_rows(reader, "Sheet1")
|> Enum.take(100)

# Always close when done
Celex.StreamReader.close(reader)
```

## API Reference

### Writing

| Function | Description |
|----------|-------------|
| `Celex.new/0` | Create a new workbook |
| `Celex.add_sheet/2` | Add a sheet to the workbook |
| `Celex.put_cell/4` | Set a cell value with optional style/formula |
| `Celex.merge_cells/2` | Merge a range of cells |
| `Celex.set_column_width/3` | Set column width |
| `Celex.set_row_height/3` | Set row height |
| `Celex.write/2` | Write workbook to file |
| `Celex.write!/2` | Write workbook to file (raises on error) |
| `Celex.stream_write/2` | Stream write with callback |

### Reading

| Function | Description |
|----------|-------------|
| `Celex.read/1` | Read workbook from file |
| `Celex.read!/1` | Read workbook from file (raises on error) |
| `Celex.get_sheet/2` | Get sheet by name or index |
| `Celex.get_cell/2` | Get cell value |
| `Celex.get_rows/1` | Get all rows from sheet |
| `Celex.sheet_names/1` | Get list of sheet names |
| `Celex.stream_open/1` | Open file for streaming read |
| `Celex.stream_sheet/2` | Stream rows from a sheet |

## Performance

Celex is designed for performance with large files:

- **Streaming write**: Rows are written directly to temporary files, then assembled into the final ZIP. Memory usage is constant regardless of file size.
- **Streaming read**: XML is parsed incrementally. Only the current row is held in memory.
- **Inline strings**: In streaming mode, strings are written inline rather than using a shared string table, trading file size for memory efficiency.

### Benchmarks

Run benchmarks with: `mix test --include benchmark`

| Operation | Rows | Duration | Memory Growth | Throughput |
|-----------|------|----------|---------------|------------|
| Write | 10,000 | 76ms | 0.01MB | ~130k rows/s |
| Write | 100,000 | 788ms | 15MB | ~127k rows/s |
| Write | 500,000 | 3.9s | 71MB | ~127k rows/s |
| Read | 10,000 | 199ms | 0.22MB | ~50k rows/s |
| Read | 100,000 | 1.95s | ~0MB | ~51k rows/s |
| Read | 500,000 | 9.8s | ~0MB | ~51k rows/s |

File size scales linearly at ~23-25 bytes per row. Reading uses SAX-based XML parsing for true O(1) memory usage.

## License

MIT License. See [LICENSE](LICENSE) for details.
