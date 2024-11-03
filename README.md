# Excel/XLSX generator in Elixir

##  Usage

### Basic

```elixir
# Create sample data
rows = [
  ["Name", "Age", "City"],
  ["John", 30, "New York"],
  ["Alice", 25, "London"],
  ["Bob", 35, "Paris"],
  ["Average age", "=AVERAGEA(B2:B4)"]
]

# Create the Excel file
Excelixir.create_excel("simple.xlsx", rows)
```

### With cell styles

# Create styled cells

```elixir
# Regular cells work as before
rows = [
  [Excelixir.Cell.new("Title", bold: true, font_size: 14), "Regular cell", Excelixir.Cell.new("Important", background_color: "FFFF00")],
  ["Plain text", 42, Excelixir.Cell.new("=SUM(A3:C3)", background_color: "FFFF00")],
  [10, 20, 30],
]

Excelixir.create_excel("styles.xlsx", rows)
```

Currently supported styles:
- `bold`
- `italic`
- `underline`
- `font_size`
- `font_color`
- `background_color`