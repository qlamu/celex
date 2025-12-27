defmodule Celex.XML.WorksheetXML do
  @moduledoc """
  Generates worksheet XML (xl/worksheets/sheetN.xml) for an XLSX package.

  This module handles the generation of cell data in various formats:
  - Inline strings (for streaming mode - no shared strings table)
  - Numbers (integers and floats)
  - Booleans
  - Dates (converted to Excel serial numbers)
  - Formulas
  """

  alias Celex.Sheet
  alias Celex.Cell
  alias Celex.Utils.{CellReference, DateUtils}

  @main_ns "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  @doc """
  Generates the complete worksheet XML from a Sheet struct.
  """
  @spec generate(Sheet.t()) :: String.t()
  def generate(%Sheet{} = sheet) do
    dimension = build_dimension(sheet)
    cols = build_columns(sheet)
    sheet_data = build_sheet_data(sheet)
    auto_filter = build_auto_filter(sheet)
    merge_cells = build_merge_cells(sheet)

    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="#{@main_ns}">
    #{dimension}
    #{cols}
    <sheetData>
    #{sheet_data}
    </sheetData>
    #{auto_filter}
    #{merge_cells}
    </worksheet>
    """
    |> String.trim()
  end

  @doc """
  Generates a single row XML from a row index and list of values/cells.

  This is useful for streaming mode where rows are written one at a time.
  """
  @spec generate_row(pos_integer(), [any()]) :: String.t()
  def generate_row(row_index, values) when is_integer(row_index) and is_list(values) do
    cells_xml =
      values
      |> Enum.with_index(1)
      |> Enum.map(fn {value, col} ->
        cell = normalize_cell(value)
        ref = CellReference.to_string({row_index, col})
        build_cell(ref, cell)
      end)
      |> Enum.join("\n")

    "<row r=\"#{row_index}\">\n#{cells_xml}\n</row>"
  end

  @doc """
  Generates just the opening worksheet and sheetData tags.
  Used for streaming writes.
  """
  @spec generate_header(keyword()) :: String.t()
  def generate_header(opts \\ []) do
    dimension = Keyword.get(opts, :dimension)
    cols = Keyword.get(opts, :columns, [])

    dim_xml =
      if dimension do
        {start_cell, end_cell} = dimension
        start_ref = CellReference.to_string(start_cell)
        end_ref = CellReference.to_string(end_cell)
        "<dimension ref=\"#{start_ref}:#{end_ref}\"/>\n"
      else
        ""
      end

    cols_xml =
      if cols != [] do
        build_columns_from_list(cols) <> "\n"
      else
        ""
      end

    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="#{@main_ns}">
    #{dim_xml}#{cols_xml}<sheetData>
    """
    |> String.trim_trailing()
  end

  @doc """
  Generates the closing tags for a worksheet.
  Used for streaming writes.
  """
  @spec generate_footer(keyword()) :: String.t()
  def generate_footer(opts \\ []) do
    merge_cells = Keyword.get(opts, :merge_cells, [])
    auto_filter = Keyword.get(opts, :auto_filter)

    auto_filter_xml =
      if auto_filter do
        "\n<autoFilter ref=\"#{auto_filter}\"/>"
      else
        ""
      end

    merge_xml =
      if merge_cells != [] do
        cells = Enum.map(merge_cells, &"<mergeCell ref=\"#{&1}\"/>") |> Enum.join("\n")
        "\n<mergeCells count=\"#{length(merge_cells)}\">\n#{cells}\n</mergeCells>"
      else
        ""
      end

    "</sheetData>#{auto_filter_xml}#{merge_xml}\n</worksheet>"
  end

  # Private functions

  defp build_dimension(%Sheet{} = sheet) do
    case Sheet.dimensions(sheet) do
      nil ->
        ""

      {start_cell, end_cell} ->
        start_ref = CellReference.to_string(start_cell)
        end_ref = CellReference.to_string(end_cell)
        "<dimension ref=\"#{start_ref}:#{end_ref}\"/>"
    end
  end

  defp build_columns(%Sheet{column_widths: widths}) when map_size(widths) == 0, do: ""

  defp build_columns(%Sheet{column_widths: widths}) do
    cols =
      widths
      |> Enum.map(fn {col, width} ->
        col_index = normalize_column_index(col)
        "<col min=\"#{col_index}\" max=\"#{col_index}\" width=\"#{width}\" customWidth=\"1\"/>"
      end)
      |> Enum.join("\n")

    "<cols>\n#{cols}\n</cols>"
  end

  defp build_columns_from_list(cols) do
    cols_xml =
      cols
      |> Enum.map(fn {col, width} ->
        col_index = normalize_column_index(col)
        "<col min=\"#{col_index}\" max=\"#{col_index}\" width=\"#{width}\" customWidth=\"1\"/>"
      end)
      |> Enum.join("\n")

    "<cols>\n#{cols_xml}\n</cols>"
  end

  defp normalize_column_index(col) when is_integer(col), do: col
  defp normalize_column_index(col) when is_binary(col), do: CellReference.column_to_index(col)

  defp build_sheet_data(%Sheet{cells: cells}) when map_size(cells) == 0, do: ""

  defp build_sheet_data(%Sheet{cells: cells}) do
    # Group cells by row
    cells
    |> Enum.group_by(fn {{row, _col}, _cell} -> row end)
    |> Enum.sort_by(fn {row, _} -> row end)
    |> Enum.map(fn {row, row_cells} ->
      # Sort cells within row by column
      sorted_cells =
        row_cells
        |> Enum.sort_by(fn {{_row, col}, _cell} -> col end)
        |> Enum.map(fn {{row, col}, cell} ->
          ref = CellReference.to_string({row, col})
          build_cell(ref, cell)
        end)
        |> Enum.join("\n")

      "<row r=\"#{row}\">\n#{sorted_cells}\n</row>"
    end)
    |> Enum.join("\n")
  end

  defp build_merge_cells(%Sheet{merged_regions: []}) do
    ""
  end

  defp build_merge_cells(%Sheet{merged_regions: regions}) do
    cells = Enum.map(regions, &"<mergeCell ref=\"#{&1}\"/>") |> Enum.join("\n")
    "<mergeCells count=\"#{length(regions)}\">\n#{cells}\n</mergeCells>"
  end

  defp build_auto_filter(%Sheet{auto_filter: nil}), do: ""

  defp build_auto_filter(%Sheet{auto_filter: range}) do
    "<autoFilter ref=\"#{range}\"/>"
  end

  defp normalize_cell(%Cell{} = cell), do: cell
  defp normalize_cell(value), do: Cell.new(value)

  defp build_cell(ref, %Cell{formula: formula} = cell) when formula != nil do
    # Formula cell
    value_xml =
      if cell.value != nil do
        "<v>#{format_value(cell.value, cell.type)}</v>"
      else
        ""
      end

    "<c r=\"#{ref}\"><f>#{escape_xml(formula)}</f>#{value_xml}</c>"
  end

  defp build_cell(ref, %Cell{type: :blank}) do
    "<c r=\"#{ref}\"/>"
  end

  defp build_cell(ref, %Cell{type: :string, value: value}) do
    # Inline string (for streaming mode - doesn't use shared strings table)
    escaped = escape_xml(to_string(value))
    "<c r=\"#{ref}\" t=\"inlineStr\"><is><t>#{escaped}</t></is></c>"
  end

  defp build_cell(ref, %Cell{type: :number, value: value}) do
    "<c r=\"#{ref}\"><v>#{value}</v></c>"
  end

  defp build_cell(ref, %Cell{type: :boolean, value: value}) do
    bool_val = if value, do: "1", else: "0"
    "<c r=\"#{ref}\" t=\"b\"><v>#{bool_val}</v></c>"
  end

  defp build_cell(ref, %Cell{type: :date, value: %Date{} = date}) do
    serial = DateUtils.date_to_serial(date)
    # s="1" could be used for date style, but we'll keep it simple for now
    "<c r=\"#{ref}\"><v>#{serial}</v></c>"
  end

  defp build_cell(ref, %Cell{type: :datetime, value: %NaiveDateTime{} = dt}) do
    serial = DateUtils.datetime_to_serial(dt)
    "<c r=\"#{ref}\"><v>#{serial}</v></c>"
  end

  defp build_cell(ref, %Cell{type: :time, value: %Time{} = time}) do
    serial = DateUtils.time_to_serial(time)
    "<c r=\"#{ref}\"><v>#{serial}</v></c>"
  end

  defp build_cell(ref, %Cell{value: value}) do
    # Fallback - treat as string
    escaped = escape_xml(to_string(value))
    "<c r=\"#{ref}\" t=\"inlineStr\"><is><t>#{escaped}</t></is></c>"
  end

  defp format_value(value, :number), do: value
  defp format_value(value, :boolean), do: if(value, do: "1", else: "0")
  defp format_value(%Date{} = date, :date), do: DateUtils.date_to_serial(date)
  defp format_value(%NaiveDateTime{} = dt, :datetime), do: DateUtils.datetime_to_serial(dt)
  defp format_value(value, _type), do: escape_xml(to_string(value))

  defp escape_xml(string) do
    string
    |> String.replace("&", "&amp;")
    |> String.replace("<", "&lt;")
    |> String.replace(">", "&gt;")
    |> String.replace("\"", "&quot;")
    |> String.replace("'", "&apos;")
  end
end
