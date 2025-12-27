defmodule Celex.Writer do
  @moduledoc """
  Non-streaming XLSX writer that takes a Workbook and writes it to a file.

  For large files or streaming data, use `Celex.StreamWriter` instead.

  ## Examples

      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Data")
        |> Workbook.update_sheet("Data", fn sheet ->
          sheet
          |> Sheet.put_cell("A1", "Hello")
          |> Sheet.put_cell("B1", "World")
        end)

      Writer.write(workbook, "output.xlsx")
  """

  alias Celex.Workbook
  alias Celex.Sheet
  alias Celex.StreamWriter

  @doc """
  Writes a workbook to the specified path.

  Returns `:ok` on success or `{:error, reason}` on failure.
  """
  @spec write(Workbook.t(), String.t()) :: :ok | {:error, term()}
  def write(%Workbook{} = workbook, path) when is_binary(path) do
    case StreamWriter.open(path) do
      {:ok, writer} ->
        try do
          writer =
            Enum.reduce(workbook.sheets, writer, fn sheet, w ->
              write_sheet(w, sheet)
            end)

          StreamWriter.close(writer)
        rescue
          e ->
            StreamWriter.abort(writer)
            {:error, Exception.message(e)}
        end
    end
  end

  @doc """
  Writes a workbook to the specified path, raising on error.
  """
  @spec write!(Workbook.t(), String.t()) :: :ok
  def write!(%Workbook{} = workbook, path) do
    case write(workbook, path) do
      :ok -> :ok
      {:error, reason} -> raise "Failed to write XLSX: #{inspect(reason)}"
    end
  end

  defp write_sheet(writer, %Sheet{} = sheet) do
    writer = StreamWriter.start_sheet(writer, sheet.name)

    # Set column widths
    writer =
      Enum.reduce(sheet.column_widths, writer, fn {col, width}, w ->
        StreamWriter.set_column_width(w, col, width)
      end)

    # Add merged regions
    writer =
      Enum.reduce(sheet.merged_regions, writer, fn region, w ->
        StreamWriter.merge_cells(w, region)
      end)

    # Write rows
    case Sheet.dimensions(sheet) do
      nil ->
        writer

      {{min_row, _min_col}, {max_row, max_col}} ->
        Enum.reduce(min_row..max_row, writer, fn row_index, w ->
          row_values =
            Enum.map(1..max_col, fn col ->
              case Sheet.get_cell(sheet, {row_index, col}) do
                nil -> nil
                cell -> cell.value
              end
            end)

          StreamWriter.write_row(w, row_values)
        end)
    end
  end
end
