defmodule Celex.Utils.CellReference do
  @moduledoc """
  Utilities for parsing and generating Excel cell references (A1 notation).

  Cell references in Excel use a combination of column letters (A-XFD) and
  row numbers (1-1048576). This module handles conversion between:

  - String references like "A1", "AA100", "XFD1048576"
  - Tuple format {row, col} where both are 1-based integers

  ## Examples

      iex> Celex.Utils.CellReference.parse("A1")
      {:ok, {1, 1}}

      iex> Celex.Utils.CellReference.to_string({1, 27})
      "AA1"

      iex> Celex.Utils.CellReference.column_to_index("AA")
      27
  """

  @doc """
  Parses a cell reference string into a {row, col} tuple.

  Returns `{:ok, {row, col}}` on success or `{:error, :invalid_reference}` on failure.
  Both row and column are 1-based.

  ## Examples

      iex> Celex.Utils.CellReference.parse("A1")
      {:ok, {1, 1}}

      iex> Celex.Utils.CellReference.parse("AA100")
      {:ok, {100, 27}}

      iex> Celex.Utils.CellReference.parse("invalid")
      {:error, :invalid_reference}
  """
  @spec parse(String.t()) :: {:ok, {pos_integer(), pos_integer()}} | {:error, :invalid_reference}
  def parse(reference) when is_binary(reference) do
    reference = String.upcase(reference)

    case Regex.run(~r/^([A-Z]+)(\d+)$/, reference) do
      [_, col_str, row_str] ->
        col = column_to_index(col_str)
        row = String.to_integer(row_str)

        if row > 0 do
          {:ok, {row, col}}
        else
          {:error, :invalid_reference}
        end

      _ ->
        {:error, :invalid_reference}
    end
  end

  @doc """
  Converts a {row, col} tuple to a cell reference string.

  ## Examples

      iex> Celex.Utils.CellReference.to_string({1, 1})
      "A1"

      iex> Celex.Utils.CellReference.to_string({100, 27})
      "AA100"
  """
  @spec to_string({pos_integer(), pos_integer()}) :: String.t()
  def to_string({row, col}) when is_integer(row) and is_integer(col) and row > 0 and col > 0 do
    "#{index_to_column(col)}#{row}"
  end

  @doc """
  Converts a column letter string to a 1-based index.

  ## Examples

      iex> Celex.Utils.CellReference.column_to_index("A")
      1

      iex> Celex.Utils.CellReference.column_to_index("Z")
      26

      iex> Celex.Utils.CellReference.column_to_index("AA")
      27
  """
  @spec column_to_index(String.t()) :: pos_integer()
  def column_to_index(column) when is_binary(column) do
    column
    |> String.upcase()
    |> String.to_charlist()
    |> Enum.reduce(0, fn char, acc ->
      acc * 26 + (char - ?A + 1)
    end)
  end

  @doc """
  Converts a 1-based column index to a column letter string.

  ## Examples

      iex> Celex.Utils.CellReference.index_to_column(1)
      "A"

      iex> Celex.Utils.CellReference.index_to_column(26)
      "Z"

      iex> Celex.Utils.CellReference.index_to_column(27)
      "AA"
  """
  @spec index_to_column(pos_integer()) :: String.t()
  def index_to_column(index) when is_integer(index) and index > 0 do
    do_index_to_column(index, [])
  end

  defp do_index_to_column(0, acc), do: List.to_string(acc)

  defp do_index_to_column(index, acc) do
    # Adjust for 1-based indexing (A=1, not A=0)
    remainder = rem(index - 1, 26)
    char = ?A + remainder
    do_index_to_column(div(index - 1, 26), [char | acc])
  end

  @doc """
  Parses a cell range string (e.g., "A1:B10") into start and end tuples.

  Returns `{:ok, {{start_row, start_col}, {end_row, end_col}}}` on success.

  ## Examples

      iex> Celex.Utils.CellReference.parse_range("A1:B10")
      {:ok, {{1, 1}, {10, 2}}}

      iex> Celex.Utils.CellReference.parse_range("AA1:AB100")
      {:ok, {{1, 27}, {100, 28}}}
  """
  @spec parse_range(String.t()) ::
          {:ok, {{pos_integer(), pos_integer()}, {pos_integer(), pos_integer()}}}
          | {:error, :invalid_range}
  def parse_range(range) when is_binary(range) do
    case String.split(range, ":") do
      [start_ref, end_ref] when start_ref != "" and end_ref != "" ->
        with {:ok, start_cell} <- parse(start_ref),
             {:ok, end_cell} <- parse(end_ref) do
          {:ok, {start_cell, end_cell}}
        else
          _ -> {:error, :invalid_range}
        end

      _ ->
        {:error, :invalid_range}
    end
  end
end
