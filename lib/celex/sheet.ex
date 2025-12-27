defmodule Celex.Sheet do
  @moduledoc """
  Represents a worksheet in an Excel workbook.

  A Sheet contains:
  - `name` - The sheet name
  - `cells` - A map of {row, col} => Cell for sparse cell storage
  - `column_widths` - Column width settings
  - `row_heights` - Row height settings
  - `merged_regions` - List of merged cell ranges

  For streaming operations, use `Celex.StreamWriter` instead.

  ## Examples

      sheet = Sheet.new("Data")
      |> Sheet.put_cell("A1", "Hello")
      |> Sheet.put_cell("B1", 42)
      |> Sheet.set_column_width("A", 20)
  """

  alias Celex.Cell
  alias Celex.Row
  alias Celex.Utils.CellReference

  @type t :: %__MODULE__{
          name: String.t(),
          index: pos_integer() | nil,
          cells: %{{pos_integer(), pos_integer()} => Cell.t()},
          rows: [[any()]] | nil,
          column_widths: %{(String.t() | pos_integer()) => number()},
          row_heights: %{pos_integer() => number()},
          merged_regions: [String.t()],
          hidden: boolean(),
          auto_filter: String.t() | nil
        }

  defstruct name: "",
            index: nil,
            cells: %{},
            rows: nil,
            column_widths: %{},
            row_heights: %{},
            merged_regions: [],
            hidden: false,
            auto_filter: nil

  @doc """
  Creates a new sheet with the given name.

  ## Options

  - `:index` - The sheet index (0-based)
  - `:hidden` - Whether the sheet is hidden

  ## Examples

      Sheet.new("Data")
      Sheet.new("Hidden", hidden: true)
  """
  @spec new(String.t(), keyword()) :: t()
  def new(name, opts \\ []) when is_binary(name) do
    %__MODULE__{
      name: name,
      index: Keyword.get(opts, :index),
      hidden: Keyword.get(opts, :hidden, false)
    }
  end

  @doc """
  Puts a cell value at the given location.

  Location can be:
  - A1 notation string: "A1", "B2", "AA100"
  - {row, col} tuple: {1, 1}, {2, 3}

  ## Options

  - `:style` - A style map for the cell
  - `:formula` - A formula string

  ## Examples

      Sheet.put_cell(sheet, "A1", "Hello")
      Sheet.put_cell(sheet, {1, 2}, 42)
      Sheet.put_cell(sheet, "A3", nil, formula: "SUM(A1:A2)")
  """
  @spec put_cell(t(), String.t() | {pos_integer(), pos_integer()}, any(), keyword()) :: t()
  def put_cell(sheet, location, value, opts \\ [])

  def put_cell(%__MODULE__{} = sheet, location, value, opts) when is_binary(location) do
    case CellReference.parse(location) do
      {:ok, {row, col}} ->
        put_cell(sheet, {row, col}, value, opts)

      {:error, _} ->
        raise ArgumentError, "Invalid cell reference: #{location}"
    end
  end

  def put_cell(%__MODULE__{cells: cells} = sheet, {row, col}, value, opts)
      when is_integer(row) and is_integer(col) do
    cell = Cell.new(value, opts)
    %{sheet | cells: Map.put(cells, {row, col}, cell)}
  end

  @doc """
  Gets the cell at the given location.

  Returns nil if no cell exists at the location.
  """
  @spec get_cell(t(), String.t() | {pos_integer(), pos_integer()}) :: Cell.t() | nil
  def get_cell(%__MODULE__{} = sheet, location) when is_binary(location) do
    case CellReference.parse(location) do
      {:ok, {row, col}} -> get_cell(sheet, {row, col})
      {:error, _} -> nil
    end
  end

  def get_cell(%__MODULE__{cells: cells}, {row, col})
      when is_integer(row) and is_integer(col) do
    Map.get(cells, {row, col})
  end

  @doc """
  Puts a row of values at the given row index.

  Values can be a list or a Row struct.
  """
  @spec put_row(t(), pos_integer(), [any()] | Row.t()) :: t()
  def put_row(%__MODULE__{} = sheet, row_index, %Row{cells: cells}) do
    cells
    |> Enum.with_index(1)
    |> Enum.reduce(sheet, fn {cell, col}, acc ->
      %{acc | cells: Map.put(acc.cells, {row_index, col}, cell)}
    end)
  end

  def put_row(%__MODULE__{} = sheet, row_index, values) when is_list(values) do
    values
    |> Enum.with_index(1)
    |> Enum.reduce(sheet, fn {value, col}, acc ->
      cell =
        case value do
          %Cell{} = c -> c
          v -> Cell.new(v)
        end

      %{acc | cells: Map.put(acc.cells, {row_index, col}, cell)}
    end)
  end

  @doc """
  Gets all cells in a row as a list.

  Returns cells from column 1 to the max column in that row.
  """
  @spec get_row(t(), pos_integer()) :: [Cell.t()]
  def get_row(%__MODULE__{cells: cells}, row_index) do
    row_cells =
      cells
      |> Enum.filter(fn {{r, _c}, _cell} -> r == row_index end)
      |> Enum.sort_by(fn {{_r, c}, _cell} -> c end)

    case row_cells do
      [] ->
        []

      _ ->
        max_col = row_cells |> Enum.map(fn {{_r, c}, _} -> c end) |> Enum.max()

        Enum.map(1..max_col, fn col ->
          Map.get(cells, {row_index, col}, Cell.new(nil))
        end)
    end
  end

  @doc """
  Sets the width for a column.

  Column can be a letter ("A") or 1-based index (1).
  """
  @spec set_column_width(t(), String.t() | pos_integer(), number()) :: t()
  def set_column_width(%__MODULE__{column_widths: widths} = sheet, column, width) do
    %{sheet | column_widths: Map.put(widths, column, width)}
  end

  @doc """
  Sets the height for a row.
  """
  @spec set_row_height(t(), pos_integer(), number()) :: t()
  def set_row_height(%__MODULE__{row_heights: heights} = sheet, row, height) do
    %{sheet | row_heights: Map.put(heights, row, height)}
  end

  @doc """
  Adds a merged cell region.

  Region should be in A1:B2 format.
  """
  @spec merge_cells(t(), String.t()) :: t()
  def merge_cells(%__MODULE__{merged_regions: regions} = sheet, range) do
    %{sheet | merged_regions: [range | regions]}
  end

  @doc """
  Sets the auto filter range for the sheet.

  The range should be in A1:B1 format, typically covering the header row.
  Pass `nil` to remove the auto filter.

  ## Examples

      Sheet.set_auto_filter(sheet, "A1:D1")
      Sheet.set_auto_filter(sheet, nil)  # Remove filter
  """
  @spec set_auto_filter(t(), String.t() | nil) :: t()
  def set_auto_filter(%__MODULE__{} = sheet, range) do
    %{sheet | auto_filter: range}
  end

  @doc """
  Returns the number of cells in the sheet.
  """
  @spec cell_count(t()) :: non_neg_integer()
  def cell_count(%__MODULE__{cells: cells}) do
    map_size(cells)
  end

  @doc """
  Returns the dimensions of the used cell range.

  Returns `{{start_row, start_col}, {end_row, end_col}}` or nil for empty sheet.
  """
  @spec dimensions(t()) :: {{pos_integer(), pos_integer()}, {pos_integer(), pos_integer()}} | nil
  def dimensions(%__MODULE__{cells: cells}) when map_size(cells) == 0, do: nil

  def dimensions(%__MODULE__{cells: cells}) do
    keys = Map.keys(cells)
    rows = Enum.map(keys, fn {r, _} -> r end)
    cols = Enum.map(keys, fn {_, c} -> c end)

    {{Enum.min(rows), Enum.min(cols)}, {Enum.max(rows), Enum.max(cols)}}
  end

  @doc """
  Returns a stream of rows in the sheet.

  Each row is a list of Cell structs from column 1 to max column.
  """
  @spec rows(t()) :: Enumerable.t()
  def rows(%__MODULE__{} = sheet) do
    case dimensions(sheet) do
      nil ->
        []

      {{min_row, _min_col}, {max_row, _max_col}} ->
        Stream.map(min_row..max_row, fn row_index ->
          get_row(sheet, row_index)
        end)
    end
  end
end
