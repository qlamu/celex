# Basic XLSX generator in Elixir

## Usage

```elixir
# Create sample data
rows = [
  ["Name", "Age", "City"],
  ["John", 30, "New York"],
  ["Alice", 25, "London"],
  ["Bob", 35, "Paris"]
]

# Create the Excel file
ExcelCreator.create_excel("output.xlsx", rows)
```