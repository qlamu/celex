defmodule Celex.Row do
  @moduledoc """
  Represents a row of cells in an Excel worksheet.

  A Row is used for streaming operations where data is processed row by row.
  Each row contains a list of cells and optional metadata like row index and height.

  ## Examples

      # Create from values
      row = Row.new(["Name", "Age", "City"])

      # Create from map with column order
      row = Row.from_map(%{name: "Alice", age: 30}, [:name, :age])

      # With options
      row = Row.new(["A", "B"], index: 1, height: 20)
  """

  alias Celex.Cell

  @type t :: %__MODULE__{
          cells: [Cell.t()],
          index: pos_integer() | nil,
          height: number() | nil,
          style: map() | nil
        }

  defstruct cells: [], index: nil, height: nil, style: nil

  @doc """
  Creates a new row from a list of values or cells.

  Values are automatically converted to Cell structs. If a value is already
  a Cell struct, it is used as-is.

  ## Options

  - `:index` - The 1-based row index in the worksheet
  - `:height` - The row height in points
  - `:style` - A style map to apply to all cells (unless cell already has style)

  ## Examples

      Row.new(["A", "B", "C"])
      Row.new([1, 2, 3], index: 1)
      Row.new(["Header1", "Header2"], style: %{bold: true})
  """
  @spec new([Cell.t() | any()], keyword()) :: t()
  def new(values, opts \\ []) when is_list(values) do
    style = Keyword.get(opts, :style)

    cells =
      Enum.map(values, fn
        %Cell{} = cell ->
          if style && cell.style == nil do
            %{cell | style: style}
          else
            cell
          end

        value ->
          if style do
            Cell.new(value, style: style)
          else
            Cell.new(value)
          end
      end)

    %__MODULE__{
      cells: cells,
      index: Keyword.get(opts, :index),
      height: Keyword.get(opts, :height),
      style: style
    }
  end

  @doc """
  Creates a row from a map, using the given column order.

  Missing keys in the map result in nil values.

  ## Examples

      data = %{name: "Alice", age: 30, city: "NYC"}
      Row.from_map(data, [:name, :age, :city])
      # => Row with cells ["Alice", 30, "NYC"]
  """
  @spec from_map(map(), [atom() | String.t()], keyword()) :: t()
  def from_map(data, columns, opts \\ []) when is_map(data) and is_list(columns) do
    values = Enum.map(columns, fn col -> Map.get(data, col) end)
    new(values, opts)
  end

  @doc """
  Extracts the values from all cells in the row.

  ## Examples

      row = Row.new(["A", 1, true])
      Row.values(row)
      # => ["A", 1, true]
  """
  @spec values(t()) :: [any()]
  def values(%__MODULE__{cells: cells}) do
    Enum.map(cells, & &1.value)
  end

  @doc """
  Returns the number of cells in the row.
  """
  @spec size(t()) :: non_neg_integer()
  def size(%__MODULE__{cells: cells}) do
    length(cells)
  end

  @doc """
  Gets the cell at the given 0-based index.

  Returns nil if the index is out of bounds.
  """
  @spec get(t(), non_neg_integer()) :: Cell.t() | nil
  def get(%__MODULE__{cells: cells}, index) when is_integer(index) and index >= 0 do
    Enum.at(cells, index)
  end

  @doc """
  Puts a value or cell at the given 0-based index.

  If the index is beyond the current row size, the row is expanded
  with blank cells to accommodate the new cell.
  """
  @spec put(t(), non_neg_integer(), Cell.t() | any()) :: t()
  def put(%__MODULE__{cells: cells} = row, index, value) when is_integer(index) and index >= 0 do
    cell =
      case value do
        %Cell{} = c -> c
        v -> Cell.new(v)
      end

    current_size = length(cells)

    updated_cells =
      if index < current_size do
        List.replace_at(cells, index, cell)
      else
        # Expand with blank cells
        padding = List.duplicate(Cell.new(nil), index - current_size)
        cells ++ padding ++ [cell]
      end

    %{row | cells: updated_cells}
  end
end
