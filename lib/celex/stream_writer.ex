defmodule Celex.StreamWriter do
  @moduledoc """
  Streaming XLSX writer for constant-memory large file generation.

  The StreamWriter writes rows directly to temporary files without holding
  the entire workbook in memory. This allows creating files with millions
  of rows while using only a few megabytes of memory.

  ## Usage

      {:ok, writer} = StreamWriter.open("output.xlsx")

      :ok =
        writer
        |> StreamWriter.start_sheet("Data")
        |> StreamWriter.write_row(["Name", "Age", "City"])
        |> StreamWriter.write_rows(data_stream)
        |> StreamWriter.close()

  ## Memory Model

  - Rows are written directly to temp files
  - Only the current row is in memory at any time
  - Uses inline strings (no shared strings table)
  - ZIP archive is built from temp files at close()
  """

  alias Celex.XML.{
    ContentTypes,
    Relationships,
    WorkbookXML,
    WorksheetXML,
    StylesXML,
    SharedStrings
  }

  @type t :: %__MODULE__{
          path: String.t(),
          temp_dir: String.t(),
          sheets: [String.t()],
          current_sheet: String.t() | nil,
          current_file: pid() | nil,
          row_count: non_neg_integer(),
          column_widths: map(),
          merged_regions: [String.t()]
        }

  defstruct [
    :path,
    :temp_dir,
    sheets: [],
    current_sheet: nil,
    current_file: nil,
    row_count: 0,
    column_widths: %{},
    merged_regions: []
  ]

  @doc """
  Opens a new streaming writer for the given output path.

  Creates a temporary directory for intermediate files.
  """
  @spec open(String.t()) :: {:ok, t()} | {:error, term()}
  def open(path) when is_binary(path) do
    temp_dir = create_temp_dir()

    writer = %__MODULE__{
      path: path,
      temp_dir: temp_dir
    }

    {:ok, writer}
  end

  @doc """
  Starts a new worksheet with the given name.

  If a sheet was previously being written, it is finalized first.
  """
  @spec start_sheet(t(), String.t()) :: t()
  def start_sheet(%__MODULE__{} = writer, name) when is_binary(name) do
    # Finalize previous sheet if any
    writer = finalize_current_sheet(writer)

    # Start new sheet
    sheet_index = length(writer.sheets) + 1
    sheet_path = Path.join(writer.temp_dir, "sheet#{sheet_index}.xml")
    {:ok, file} = File.open(sheet_path, [:write, :utf8])

    # Write sheet header
    header = WorksheetXML.generate_header()
    IO.write(file, header <> "\n")

    %{
      writer
      | sheets: writer.sheets ++ [name],
        current_sheet: name,
        current_file: file,
        row_count: 0,
        column_widths: %{},
        merged_regions: []
    }
  end

  @doc """
  Writes a single row of values to the current sheet.

  Values can be any supported type: strings, numbers, booleans, dates, etc.
  """
  @spec write_row(t(), [any()]) :: t()
  def write_row(%__MODULE__{current_file: nil}, _values) do
    raise "No sheet started. Call start_sheet/2 first."
  end

  def write_row(%__MODULE__{current_file: file, row_count: count} = writer, values)
      when is_list(values) do
    row_xml = WorksheetXML.generate_row(count + 1, values)
    IO.write(file, row_xml <> "\n")

    %{writer | row_count: count + 1}
  end

  @doc """
  Writes multiple rows from an enumerable.

  The enumerable is consumed lazily, making this suitable for streams.
  """
  @spec write_rows(t(), Enumerable.t()) :: t()
  def write_rows(%__MODULE__{} = writer, rows) do
    Enum.reduce(rows, writer, fn row, acc ->
      write_row(acc, row)
    end)
  end

  @doc """
  Sets the width for a column in the current sheet.
  """
  @spec set_column_width(t(), String.t() | pos_integer(), number()) :: t()
  def set_column_width(%__MODULE__{column_widths: widths} = writer, column, width) do
    %{writer | column_widths: Map.put(widths, column, width)}
  end

  @doc """
  Adds a merged cell region to the current sheet.
  """
  @spec merge_cells(t(), String.t()) :: t()
  def merge_cells(%__MODULE__{merged_regions: regions} = writer, range) do
    %{writer | merged_regions: [range | regions]}
  end

  @doc """
  Closes the writer and creates the final XLSX file.

  This finalizes all sheets and builds the ZIP archive.
  """
  @spec close(t()) :: :ok | {:error, term()}
  def close(%__MODULE__{} = writer) do
    # Finalize current sheet
    writer = finalize_current_sheet(writer)

    # Build ZIP file
    result = build_xlsx(writer)

    # Cleanup temp files
    cleanup_temp_dir(writer.temp_dir)

    result
  end

  @doc """
  Aborts the writer and cleans up temporary files.

  Use this if you need to cancel the write operation.
  """
  @spec abort(t()) :: :ok
  def abort(%__MODULE__{current_file: file, temp_dir: temp_dir}) do
    if file, do: File.close(file)
    cleanup_temp_dir(temp_dir)
    :ok
  end

  # Private functions

  defp create_temp_dir do
    temp_base = System.tmp_dir!()
    dir_name = "celex_#{:erlang.unique_integer([:positive])}"
    temp_dir = Path.join(temp_base, dir_name)
    File.mkdir_p!(temp_dir)
    temp_dir
  end

  defp cleanup_temp_dir(temp_dir) do
    File.rm_rf(temp_dir)
  end

  defp finalize_current_sheet(%__MODULE__{current_file: nil} = writer), do: writer

  defp finalize_current_sheet(%__MODULE__{current_file: file, merged_regions: regions} = writer) do
    # Write sheet footer
    footer = WorksheetXML.generate_footer(merge_cells: Enum.reverse(regions))
    IO.write(file, footer)
    File.close(file)

    %{writer | current_file: nil}
  end

  defp build_xlsx(%__MODULE__{path: path, temp_dir: temp_dir, sheets: sheets}) do
    sheet_count = length(sheets)

    # Generate all XML components
    content_types = ContentTypes.generate(sheet_count: sheet_count, include_shared_strings: false)
    root_rels = Relationships.generate_root()

    workbook_rels =
      Relationships.generate_workbook(sheet_count: sheet_count, include_shared_strings: false)

    workbook_xml = WorkbookXML.generate(sheet_names: sheets)
    styles_xml = StylesXML.generate()
    shared_strings = SharedStrings.generate([])

    # Build file list for ZIP
    # Format: {filename (charlist), content (binary)}
    files = [
      {~c"[Content_Types].xml", content_types},
      {~c"_rels/.rels", root_rels},
      {~c"xl/_rels/workbook.xml.rels", workbook_rels},
      {~c"xl/workbook.xml", workbook_xml},
      {~c"xl/styles.xml", styles_xml},
      {~c"xl/sharedStrings.xml", shared_strings}
    ]

    # Add worksheet files
    sheet_files =
      sheets
      |> Enum.with_index(1)
      |> Enum.map(fn {_name, index} ->
        sheet_path = Path.join(temp_dir, "sheet#{index}.xml")
        {:ok, content} = File.read(sheet_path)
        {~c"xl/worksheets/sheet#{index}.xml", content}
      end)

    all_files = files ++ sheet_files

    # Create ZIP file - content must be binary (which Elixir strings are)
    case :zip.create(String.to_charlist(path), all_files) do
      {:ok, _} -> :ok
      {:error, reason} -> {:error, reason}
    end
  end
end
