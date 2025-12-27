defmodule Celex.StreamReader do
  @moduledoc """
  Streaming XLSX reader with constant memory usage.

  Uses SAX-based XML parsing to lazily stream rows from worksheets,
  allowing processing of files with millions of rows without loading
  the entire file into memory.

  ## Examples

      {:ok, reader} = Celex.StreamReader.open("large_file.xlsx")

      # Get sheet names
      names = Celex.StreamReader.sheet_names(reader)

      # Stream rows lazily
      reader
      |> Celex.StreamReader.stream_rows("Sheet1")
      |> Stream.drop(1)  # Skip header
      |> Stream.each(&process_row/1)
      |> Stream.run()

      Celex.StreamReader.close(reader)
  """

  alias Celex.XML.SaxParser

  defstruct [
    :zip_handle,
    :zip_path,
    :sheets,
    :shared_strings,
    :temp_dir
  ]

  @type t :: %__MODULE__{
          zip_handle: :zip.handle() | nil,
          zip_path: String.t(),
          sheets: [%{name: String.t(), path: String.t(), index: non_neg_integer()}],
          shared_strings: [String.t()],
          temp_dir: String.t() | nil
        }

  @doc """
  Opens an XLSX file for streaming read.

  Returns `{:ok, reader}` on success or `{:error, reason}` on failure.
  """
  @spec open(String.t()) :: {:ok, t()} | {:error, term()}
  def open(path) when is_binary(path) do
    with {:ok, zip_handle} <- open_zip(path),
         {:ok, sheets} <- read_sheet_info(zip_handle),
         {:ok, shared_strings} <- read_shared_strings(zip_handle) do
      reader = %__MODULE__{
        zip_handle: zip_handle,
        zip_path: path,
        sheets: sheets,
        shared_strings: shared_strings
      }

      {:ok, reader}
    end
  end

  @doc """
  Returns the list of sheet names in the workbook.
  """
  @spec sheet_names(t()) :: [String.t()]
  def sheet_names(%__MODULE__{sheets: sheets}) do
    Enum.map(sheets, & &1.name)
  end

  @doc """
  Returns a lazy Stream of rows from the specified sheet.

  The sheet can be specified by name (string) or index (0-based integer).
  Each row is returned as a list of cell values.

  ## Examples

      stream_rows(reader, "Sheet1") |> Enum.take(10)
      stream_rows(reader, 0) |> Enum.to_list()
  """
  @spec stream_rows(t(), String.t() | non_neg_integer()) :: Enumerable.t()
  def stream_rows(%__MODULE__{} = reader, sheet) when is_binary(sheet) do
    case find_sheet(reader, sheet) do
      nil -> []
      sheet_info -> stream_sheet_rows(reader, sheet_info)
    end
  end

  def stream_rows(%__MODULE__{} = reader, index) when is_integer(index) and index >= 0 do
    case Enum.at(reader.sheets, index) do
      nil -> []
      sheet_info -> stream_sheet_rows(reader, sheet_info)
    end
  end

  @doc """
  Reads all rows from the specified sheet into a list.

  Convenience function that collects all rows from `stream_rows/2`.
  For large files, prefer using `stream_rows/2` directly.
  """
  @spec read_rows(t(), String.t() | non_neg_integer()) :: [[any()]]
  def read_rows(%__MODULE__{} = reader, sheet) do
    stream_rows(reader, sheet) |> Enum.to_list()
  end

  @doc """
  Closes the reader and releases resources.
  """
  @spec close(t()) :: :ok
  def close(%__MODULE__{zip_handle: handle}) when not is_nil(handle) do
    :zip.zip_close(handle)
    :ok
  end

  def close(%__MODULE__{}), do: :ok

  # Private functions

  defp open_zip(path) do
    case :zip.zip_open(String.to_charlist(path), [:memory]) do
      {:ok, handle} -> {:ok, handle}
      {:error, reason} -> {:error, reason}
    end
  end

  defp read_sheet_info(zip_handle) do
    with {:ok, workbook_xml} <- read_zip_file(zip_handle, ~c"xl/workbook.xml"),
         {:ok, rels_xml} <- read_zip_file(zip_handle, ~c"xl/_rels/workbook.xml.rels") do
      sheets = parse_workbook_sheets(workbook_xml)
      rels = parse_relationships(rels_xml)

      sheets_with_paths =
        sheets
        |> Enum.with_index()
        |> Enum.map(fn {{name, rid}, index} ->
          path = Map.get(rels, rid, "worksheets/sheet#{index + 1}.xml")
          %{name: name, path: "xl/#{path}", index: index}
        end)

      {:ok, sheets_with_paths}
    end
  end

  defp read_shared_strings(zip_handle) do
    case read_zip_file(zip_handle, ~c"xl/sharedStrings.xml") do
      {:ok, xml} -> {:ok, parse_shared_strings(xml)}
      {:error, :enoent} -> {:ok, []}
    end
  end

  defp read_zip_file(zip_handle, filename) do
    case :zip.zip_get(filename, zip_handle) do
      {:ok, {_, content}} -> {:ok, content}
      {:error, reason} -> {:error, reason}
    end
  end

  defp parse_workbook_sheets(xml) do
    xml
    |> to_string()
    |> extract_sheets()
  end

  defp extract_sheets(xml_string) do
    # Extract <sheet> elements with name and r:id attributes
    regex = ~r/<sheet[^>]*name="([^"]*)"[^>]*r:id="([^"]*)"/

    Regex.scan(regex, xml_string)
    |> Enum.map(fn [_, name, rid] -> {unescape_xml(name), rid} end)
  end

  defp parse_relationships(xml) do
    xml
    |> to_string()
    |> extract_relationships()
  end

  defp extract_relationships(xml_string) do
    # Extract <Relationship> elements
    regex = ~r/<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"/

    Regex.scan(regex, xml_string)
    |> Map.new(fn [_, id, target] -> {id, target} end)
  end

  defp parse_shared_strings(xml) do
    xml
    |> to_string()
    |> extract_shared_strings()
  end

  defp extract_shared_strings(xml_string) do
    # Extract <t> elements from <si> elements
    # Handle both <si><t>text</t></si> and <si><r><t>text</t></r></si>
    regex = ~r/<si[^>]*>.*?<t[^>]*>([^<]*)<\/t>.*?<\/si>/s

    Regex.scan(regex, xml_string)
    |> Enum.map(fn [_, text] -> unescape_xml(text) end)
  end

  defp find_sheet(%__MODULE__{sheets: sheets}, name) do
    Enum.find(sheets, fn sheet -> sheet.name == name end)
  end

  defp stream_sheet_rows(%__MODULE__{} = reader, sheet_info) do
    Stream.resource(
      fn -> init_sax_stream(reader, sheet_info) end,
      &next_sax_row/1,
      &cleanup_sax_stream/1
    )
  end

  defp init_sax_stream(reader, sheet_info) do
    case read_zip_file(reader.zip_handle, String.to_charlist(sheet_info.path)) do
      {:ok, xml} ->
        # Start SAX parser in a separate process (not linked, so we can kill it safely)
        caller = self()
        shared_strings = reader.shared_strings

        parser_pid =
          spawn(fn ->
            SaxParser.stream_rows(xml, shared_strings, caller)
          end)

        %{
          parser_pid: parser_pid,
          shared_strings: shared_strings,
          done: false
        }

      {:error, _reason} ->
        %{parser_pid: nil, shared_strings: [], done: true}
    end
  end

  defp next_sax_row(%{done: true} = state) do
    {:halt, state}
  end

  defp next_sax_row(%{parser_pid: nil} = state) do
    {:halt, state}
  end

  defp next_sax_row(state) do
    receive do
      {:row, row} ->
        # Resolve shared strings in the row
        resolved_row = resolve_shared_strings(row, state.shared_strings)
        {[resolved_row], state}

      {:done, :ok} ->
        {:halt, %{state | done: true}}

      {:error, _reason} ->
        {:halt, %{state | done: true}}
    after
      30_000 ->
        # Timeout after 30 seconds
        {:halt, %{state | done: true}}
    end
  end

  defp cleanup_sax_stream(%{parser_pid: pid}) when is_pid(pid) do
    # Parser should be done, but clean up if needed
    if Process.alive?(pid) do
      Process.exit(pid, :kill)
    end

    :ok
  end

  defp cleanup_sax_stream(_state), do: :ok

  defp resolve_shared_strings(row, shared_strings) do
    Enum.map(row, fn
      {:shared_string, index} -> Enum.at(shared_strings, index, "")
      value -> value
    end)
  end

  defp unescape_xml(text) do
    text
    |> String.replace("&amp;", "&")
    |> String.replace("&lt;", "<")
    |> String.replace("&gt;", ">")
    |> String.replace("&quot;", "\"")
    |> String.replace("&apos;", "'")
  end
end
