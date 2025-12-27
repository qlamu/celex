defmodule Celex.Reader do
  @moduledoc """
  Non-streaming XLSX reader that loads entire workbooks into memory.

  Best suited for small to medium-sized files. For large files (millions of rows),
  use `Celex.StreamReader` instead.

  ## Examples

      {:ok, workbook} = Celex.Reader.read("data.xlsx")

      # Get sheet by name or index
      {:ok, sheet} = Celex.Reader.get_sheet(workbook, "Sales")
      {:ok, sheet} = Celex.Reader.get_sheet(workbook, 0)

      # Get cell values
      value = Celex.Reader.get_cell(sheet, "A1")
      value = Celex.Reader.get_cell(sheet, {1, 1})

      # Get all rows
      rows = Celex.Reader.get_rows(sheet)
  """

  alias Celex.{Workbook, Sheet, StreamReader}
  alias Celex.Utils.CellReference

  @doc """
  Reads an XLSX file and returns a Workbook struct.

  Returns `{:ok, workbook}` on success or `{:error, reason}` on failure.
  """
  @spec read(String.t()) :: {:ok, Workbook.t()} | {:error, term()}
  def read(path) when is_binary(path) do
    with {:ok, reader} <- StreamReader.open(path) do
      workbook = build_workbook(reader)
      StreamReader.close(reader)
      {:ok, workbook}
    end
  end

  @doc """
  Reads an XLSX file and returns a Workbook struct, raising on error.
  """
  @spec read!(String.t()) :: Workbook.t()
  def read!(path) when is_binary(path) do
    case read(path) do
      {:ok, workbook} -> workbook
      {:error, reason} -> raise File.Error, reason: reason, action: "read", path: path
    end
  end

  @doc """
  Gets a sheet from a workbook by name or index.

  Returns `{:ok, sheet}` or `{:error, :sheet_not_found}`.
  """
  @spec get_sheet(Workbook.t(), String.t() | non_neg_integer()) ::
          {:ok, Sheet.t()} | {:error, :sheet_not_found}
  def get_sheet(%Workbook{sheets: sheets}, name) when is_binary(name) do
    case Enum.find(sheets, fn sheet -> sheet.name == name end) do
      nil -> {:error, :sheet_not_found}
      sheet -> {:ok, sheet}
    end
  end

  def get_sheet(%Workbook{sheets: sheets}, index) when is_integer(index) and index >= 0 do
    case Enum.at(sheets, index) do
      nil -> {:error, :sheet_not_found}
      sheet -> {:ok, sheet}
    end
  end

  @doc """
  Gets a cell value from a sheet by reference.

  The reference can be:
  - A string like "A1", "B2", "AA100"
  - A tuple `{row, column}` where both are 1-based integers

  Returns the cell value or `nil` if the cell is empty.
  """
  @spec get_cell(Sheet.t(), String.t() | {pos_integer(), pos_integer()}) :: any()
  def get_cell(%Sheet{} = sheet, ref) when is_binary(ref) do
    case CellReference.parse(ref) do
      {:ok, {row, col}} -> get_cell_value(sheet, row, col)
      {:error, _} -> nil
    end
  end

  def get_cell(%Sheet{} = sheet, {row, col})
      when is_integer(row) and is_integer(col) and row > 0 and col > 0 do
    get_cell_value(sheet, row, col)
  end

  @doc """
  Returns all rows from a sheet as a list of lists.

  Each row is a list of cell values, with `nil` for empty cells.
  """
  @spec get_rows(Sheet.t()) :: [[any()]]
  def get_rows(%Sheet{rows: rows}) when is_list(rows) do
    rows
  end

  def get_rows(%Sheet{cells: cells}) when is_map(cells) do
    if map_size(cells) == 0 do
      []
    else
      # Find dimensions
      {max_row, max_col} = find_dimensions(cells)

      # Build rows
      for row <- 1..max_row do
        for col <- 1..max_col do
          Map.get(cells, {row, col})
        end
      end
    end
  end

  # Private functions

  defp build_workbook(reader) do
    sheet_names = StreamReader.sheet_names(reader)

    sheets =
      Enum.map(sheet_names, fn name ->
        rows = StreamReader.read_rows(reader, name)
        build_sheet(name, rows)
      end)

    %Workbook{sheets: sheets}
  end

  defp build_sheet(name, rows) do
    # Convert rows to cells map for efficient random access
    cells =
      rows
      |> Enum.with_index(1)
      |> Enum.flat_map(fn {row, row_idx} ->
        row
        |> Enum.with_index(1)
        |> Enum.map(fn {value, col_idx} ->
          {{row_idx, col_idx}, value}
        end)
      end)
      |> Map.new()

    %Sheet{name: name, cells: cells, rows: rows}
  end

  defp get_cell_value(%Sheet{cells: cells}, row, col) when is_map(cells) do
    Map.get(cells, {row, col})
  end

  defp get_cell_value(%Sheet{rows: rows}, row, col) when is_list(rows) do
    case Enum.at(rows, row - 1) do
      nil -> nil
      row_data -> Enum.at(row_data, col - 1)
    end
  end

  defp find_dimensions(cells) do
    Enum.reduce(cells, {0, 0}, fn {{row, col}, _value}, {max_row, max_col} ->
      {max(row, max_row), max(col, max_col)}
    end)
  end
end
