defmodule Celex do
  @moduledoc """
  Celex - A robust, stream-based Elixir library for reading and writing XLSX files.

  Designed to handle large files (millions of rows) with constant memory usage.

  ## Features

  - Stream-based reading and writing for large files
  - Builder pattern API for convenience
  - Full XLSX support: styles, formulas, merged cells
  - Zero runtime dependencies (uses Erlang's built-in :zip and :xmerl)

  ## Quick Start

  ### Streaming API (for large files)

      # Writing
      Celex.stream_write("report.xlsx", fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["Name", "Age"])
        |> Celex.StreamWriter.write_rows(data_stream)
      end)

      # Reading
      "report.xlsx"
      |> Celex.stream_sheet("Data")
      |> Stream.drop(1)  # Skip header
      |> Stream.each(&process_row/1)
      |> Stream.run()

  ### Builder API (for small files)

      # Writing
      Celex.new()
      |> Celex.add_sheet("Sales")
      |> Celex.put_cell("A1", "Product")
      |> Celex.put_cell("B1", 42)
      |> Celex.write!("report.xlsx")

      # Reading
      workbook = Celex.read!("report.xlsx")
      {:ok, sheet} = Celex.get_sheet(workbook, "Sales")
      rows = Celex.get_rows(sheet)

  """

  alias Celex.{Workbook, Sheet, Cell, Writer, StreamWriter, StreamReader, Reader}

  # ============================================================================
  # Workbook creation and management
  # ============================================================================

  @doc """
  Creates a new empty workbook.

  ## Examples

      workbook = Celex.new()
  """
  @spec new() :: Workbook.t()
  def new, do: Workbook.new()

  @doc """
  Adds a new sheet to the workbook.

  ## Examples

      workbook = Celex.new() |> Celex.add_sheet("Data")
  """
  @spec add_sheet(Workbook.t(), String.t()) :: Workbook.t()
  def add_sheet(%Workbook{} = workbook, name) when is_binary(name) do
    Workbook.add_sheet(workbook, name)
  end

  @doc """
  Puts a cell value at the given location in the current (last) sheet.

  ## Options

  - `:style` - Style map for the cell
  - `:formula` - Formula string (without leading '=')

  ## Examples

      workbook
      |> Celex.put_cell("A1", "Hello")
      |> Celex.put_cell("B1", 42, style: %{bold: true})
      |> Celex.put_cell("C1", nil, formula: "SUM(A1:B1)")
  """
  @spec put_cell(Workbook.t(), String.t(), any(), keyword()) :: Workbook.t()
  def put_cell(workbook, location, value, opts \\ [])

  def put_cell(%Workbook{sheets: []} = _workbook, _location, _value, _opts) do
    raise "No sheet exists. Call add_sheet/2 first."
  end

  def put_cell(%Workbook{sheets: sheets} = workbook, location, value, opts) do
    last_sheet = List.last(sheets)

    Workbook.update_sheet(workbook, last_sheet.name, fn sheet ->
      Sheet.put_cell(sheet, location, value, opts)
    end)
  end

  @doc """
  Gets a cell from the specified sheet.

  ## Examples

      cell = Celex.get_cell(workbook, "Sheet1", "A1")
  """
  @spec get_cell(Workbook.t(), String.t(), String.t()) :: Cell.t() | nil
  def get_cell(%Workbook{} = workbook, sheet_name, location) do
    case Workbook.get_sheet(workbook, sheet_name) do
      nil -> nil
      sheet -> Sheet.get_cell(sheet, location)
    end
  end

  @doc """
  Sets the column width for a column in the current sheet.

  ## Examples

      workbook |> Celex.set_column_width("A", 20)
  """
  @spec set_column_width(Workbook.t(), String.t() | pos_integer(), number()) :: Workbook.t()
  def set_column_width(%Workbook{sheets: []} = _workbook, _column, _width) do
    raise "No sheet exists. Call add_sheet/2 first."
  end

  def set_column_width(%Workbook{sheets: sheets} = workbook, column, width) do
    last_sheet = List.last(sheets)

    Workbook.update_sheet(workbook, last_sheet.name, fn sheet ->
      Sheet.set_column_width(sheet, column, width)
    end)
  end

  @doc """
  Sets the row height for a row in the current sheet.

  ## Examples

      workbook |> Celex.set_row_height(1, 25)
  """
  @spec set_row_height(Workbook.t(), pos_integer(), number()) :: Workbook.t()
  def set_row_height(%Workbook{sheets: []} = _workbook, _row, _height) do
    raise "No sheet exists. Call add_sheet/2 first."
  end

  def set_row_height(%Workbook{sheets: sheets} = workbook, row, height) do
    last_sheet = List.last(sheets)

    Workbook.update_sheet(workbook, last_sheet.name, fn sheet ->
      Sheet.set_row_height(sheet, row, height)
    end)
  end

  @doc """
  Merges cells in the current sheet.

  ## Examples

      workbook |> Celex.merge_cells("A1:B2")
  """
  @spec merge_cells(Workbook.t(), String.t()) :: Workbook.t()
  def merge_cells(%Workbook{sheets: []} = _workbook, _range) do
    raise "No sheet exists. Call add_sheet/2 first."
  end

  def merge_cells(%Workbook{sheets: sheets} = workbook, range) do
    last_sheet = List.last(sheets)

    Workbook.update_sheet(workbook, last_sheet.name, fn sheet ->
      Sheet.merge_cells(sheet, range)
    end)
  end

  @doc """
  Sets the auto filter range for the current sheet.

  The range should be in A1:B1 format, typically covering the header row.
  This adds filter dropdowns to the specified columns in Excel.

  ## Examples

      workbook
      |> Celex.add_sheet("Data")
      |> Celex.put_cell("A1", "Name")
      |> Celex.put_cell("B1", "Age")
      |> Celex.set_auto_filter("A1:B1")
  """
  @spec set_auto_filter(Workbook.t(), String.t()) :: Workbook.t()
  def set_auto_filter(%Workbook{sheets: []} = _workbook, _range) do
    raise "No sheet exists. Call add_sheet/2 first."
  end

  def set_auto_filter(%Workbook{sheets: sheets} = workbook, range) do
    last_sheet = List.last(sheets)

    Workbook.update_sheet(workbook, last_sheet.name, fn sheet ->
      Sheet.set_auto_filter(sheet, range)
    end)
  end

  @doc """
  Returns the list of sheet names.

  ## Examples

      names = Celex.sheet_names(workbook)
      # => ["Sheet1", "Sheet2"]
  """
  @spec sheet_names(Workbook.t()) :: [String.t()]
  def sheet_names(%Workbook{} = workbook) do
    Workbook.sheet_names(workbook)
  end

  # ============================================================================
  # Writing files
  # ============================================================================

  @doc """
  Writes the workbook to a file.

  Returns `:ok` on success or `{:error, reason}` on failure.

  ## Examples

      Celex.write(workbook, "output.xlsx")
  """
  @spec write(Workbook.t(), String.t()) :: :ok | {:error, term()}
  def write(%Workbook{} = workbook, path) when is_binary(path) do
    Writer.write(workbook, path)
  end

  @doc """
  Writes the workbook to a file, raising on error.

  ## Examples

      Celex.write!(workbook, "output.xlsx")
  """
  @spec write!(Workbook.t(), String.t()) :: :ok
  def write!(%Workbook{} = workbook, path) when is_binary(path) do
    Writer.write!(workbook, path)
  end

  # ============================================================================
  # Reading API
  # ============================================================================

  @doc """
  Reads an XLSX file into a workbook struct.

  Returns `{:ok, workbook}` on success or `{:error, reason}` on failure.

  ## Examples

      {:ok, workbook} = Celex.read("data.xlsx")
      sheet_names = Celex.sheet_names(workbook)
  """
  @spec read(String.t()) :: {:ok, Workbook.t()} | {:error, term()}
  def read(path) when is_binary(path) do
    Reader.read(path)
  end

  @doc """
  Reads an XLSX file into a workbook struct, raising on error.

  ## Examples

      workbook = Celex.read!("data.xlsx")
  """
  @spec read!(String.t()) :: Workbook.t()
  def read!(path) when is_binary(path) do
    Reader.read!(path)
  end

  @doc """
  Gets a sheet from a workbook by name or index.

  Returns `{:ok, sheet}` or `{:error, :sheet_not_found}`.

  ## Examples

      {:ok, sheet} = Celex.get_sheet(workbook, "Sales")
      {:ok, sheet} = Celex.get_sheet(workbook, 0)
  """
  @spec get_sheet(Workbook.t(), String.t() | non_neg_integer()) ::
          {:ok, Sheet.t()} | {:error, :sheet_not_found}
  def get_sheet(%Workbook{} = workbook, sheet) do
    Reader.get_sheet(workbook, sheet)
  end

  @doc """
  Gets a cell value from a sheet.

  The reference can be a string like "A1" or a tuple `{row, col}`.

  ## Examples

      value = Celex.get_cell(sheet, "A1")
      value = Celex.get_cell(sheet, {1, 1})
  """
  @spec get_cell(Sheet.t(), String.t() | {pos_integer(), pos_integer()}) :: any()
  def get_cell(%Sheet{} = sheet, ref) do
    Reader.get_cell(sheet, ref)
  end

  @doc """
  Returns all rows from a sheet as a list of lists.

  ## Examples

      rows = Celex.get_rows(sheet)
      # => [["Name", "Age"], ["Alice", 30], ["Bob", 25]]
  """
  @spec get_rows(Sheet.t()) :: [[any()]]
  def get_rows(%Sheet{} = sheet) do
    Reader.get_rows(sheet)
  end

  # ============================================================================
  # Streaming API
  # ============================================================================

  @doc """
  Opens a streaming writer for writing large files with constant memory.

  The callback receives the writer and should return the final writer state.

  ## Examples

      Celex.stream_write("large_file.xlsx", fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["Header1", "Header2"])
        |> Celex.StreamWriter.write_rows(data_stream)
      end)
  """
  @spec stream_write(String.t(), (StreamWriter.t() -> StreamWriter.t())) :: :ok | {:error, term()}
  def stream_write(path, callback) when is_binary(path) and is_function(callback, 1) do
    case StreamWriter.open(path) do
      {:ok, writer} ->
        try do
          writer = callback.(writer)
          StreamWriter.close(writer)
        rescue
          e ->
            StreamWriter.abort(writer)
            {:error, Exception.message(e)}
        end
    end
  end

  @doc """
  Opens an XLSX file for streaming read with constant memory usage.

  Returns `{:ok, reader}` on success or `{:error, reason}` on failure.

  ## Examples

      {:ok, reader} = Celex.stream_open("large_file.xlsx")
      sheets = Celex.StreamReader.sheet_names(reader)

      reader
      |> Celex.StreamReader.stream_rows("Sheet1")
      |> Stream.drop(1)
      |> Stream.each(&process_row/1)
      |> Stream.run()

      Celex.StreamReader.close(reader)
  """
  @spec stream_open(String.t()) :: {:ok, StreamReader.t()} | {:error, term()}
  def stream_open(path) when is_binary(path) do
    StreamReader.open(path)
  end

  @doc """
  Streams rows from a specific sheet in an XLSX file.

  Returns a lazy Stream of rows that can be processed with constant memory.
  Each row is a list of cell values.

  ## Examples

      "data.xlsx"
      |> Celex.stream_sheet("Sheet1")
      |> Stream.drop(1)  # Skip header
      |> Enum.each(&process_row/1)

      # Stream by sheet index
      Celex.stream_sheet("data.xlsx", 0) |> Enum.take(100)
  """
  @spec stream_sheet(String.t(), String.t() | non_neg_integer()) :: Enumerable.t()
  def stream_sheet(path, sheet) when is_binary(path) do
    case StreamReader.open(path) do
      {:ok, reader} ->
        Stream.resource(
          fn -> {reader, StreamReader.read_rows(reader, sheet)} end,
          fn
            {reader, []} -> {:halt, reader}
            {reader, [row | rest]} -> {[row], {reader, rest}}
          end,
          fn reader -> StreamReader.close(reader) end
        )

      {:error, _reason} ->
        Stream.iterate(nil, fn _ -> nil end) |> Stream.take(0)
    end
  end
end
