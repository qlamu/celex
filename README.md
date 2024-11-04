# Celex

Celex is a pure Elixir library for creating Excel XLSX files without any external dependencies. It supports multiple worksheets, cell styling, and formulas while maintaining a simple, flexible API.

## Features

- No external dependencies
- Multiple worksheets support
- Cell styling (bold, italic, underline, font size, colors)
- Formula support
- Flexible input formats
- Pure Elixir implementation
- Based on the Office Open XML format

## Installation

Add celex to your list of dependencies in mix.exs:

```elixir
def deps do
  [
    {:celex, git: "https://github.com/qlamu/celex", tag: "0.1.0"}
  ]
end
```

## Usage

### Simple Single Worksheet

For basic usage, just pass a list of lists:

```elixir
Celex.create_excel("simple.xlsx", [
  ["Name", "Age", "City"],
  ["John", 30, "New York"],
  ["Alice", 25, "London"]
])
```

### Multiple Worksheets

Create multiple worksheets using a map:

```elixir
Celex.create_excel("multi.xlsx", %{
  "Sales" => [
    ["Product", "Amount"],
    ["A", 100],
    ["B", 200]
  ],
  "Costs" => [
    ["Product", "Cost"],
    ["A", 50],
    ["B", 80]
  ]
})
```

### Styled Worksheets

For more control over styling and formatting, use the structured approach:

```elixir
alias Celex.{Worksheet, Cell}

worksheets = [
  Worksheet.new("Sales", [
    [
      Cell.new("Product", bold: true, background_color: "FF4F81BD"),
      Cell.new("Q1", bold: true, background_color: "FF4F81BD"),
      Cell.new("Q2", bold: true, background_color: "FF4F81BD")
    ],
    ["Laptops", 150_000, 180_000],
    ["Phones", 200_000, 185_000],
    [
      Cell.new("Total", bold: true),
      Cell.new("=SUM(B2:B3)"),
      Cell.new("=SUM(C2:C3)")
    ]
  ]),
  
  Worksheet.new("Summary", [
    [Cell.new("Key Metrics", bold: true, font_size: 14)],
    ["Total Q2 Sales", "=Sales!C4"],
    ["Average Sales", "=AVERAGE(Sales!B2:C3)"]
  ])
]

Celex.create_excel("report.xlsx", worksheets)
```

## Styling Options

The `Cell.new/2` function supports the following styling options:

```elixir
Cell.new(value, [
  bold: true | false,
  italic: true | false,
  underline: true | false,
  font_size: number,
  font_color: "FFRRGGBB",
  background_color: "FFRRGGBB",
  number_format: "currency" | "long_time" | "scientific" | ...,
  alignment: "center" | "left" | "right"
])
```

Examples:

```elixir
# Bold red text
Cell.new("Important", bold: true, font_color: "FFFF0000")

# Yellow background
Cell.new("Highlighted", background_color: "FFFFFF00")

# Multiple styles
Cell.new("Header", [
  bold: true,
  font_size: 14,
  font_color: "FFFFFFFF",
  background_color: "FF4F81BD"
])

# Date
Cell.new(Date.utc_today(), number_format: "short_date")
```

## Formula Support

Excel formulas are supported by prefixing the cell value with "=":

```elixir
# Direct formula
Cell.new("=SUM(A1:A10)")

# Formula referencing other sheets
Cell.new("=AVERAGE(Sales!B2:B5)")
```