defmodule Celex.XML.SharedStrings do
  @moduledoc """
  Generates the xl/sharedStrings.xml file for an XLSX package.

  The shared strings table is used to deduplicate string values across the workbook.
  Each unique string is stored once and referenced by index from cells.

  Note: For streaming mode, inline strings are used instead to avoid
  needing to know all strings upfront.
  """

  @main_ns "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  @doc """
  Generates the sharedStrings.xml content from a list of unique strings.
  """
  @spec generate([String.t()]) :: String.t()
  def generate(strings) when is_list(strings) do
    count = length(strings)

    items =
      strings
      |> Enum.map(&build_string_item/1)
      |> Enum.join("\n")

    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst xmlns="#{@main_ns}" count="#{count}" uniqueCount="#{count}">
    #{items}
    </sst>
    """
    |> String.trim()
  end

  @doc """
  Builds a shared strings table from a list of strings (possibly with duplicates).

  Returns a tuple of:
  - A map of string -> index for lookup
  - A list of unique strings in order

  ## Examples

      {table, strings} = SharedStrings.build_table(["A", "B", "A", "C"])
      # table = %{"A" => 0, "B" => 1, "C" => 2}
      # strings = ["A", "B", "C"]
  """
  @spec build_table([String.t()]) :: {%{String.t() => non_neg_integer()}, [String.t()]}
  def build_table(strings) when is_list(strings) do
    {table, unique_strings, _} =
      Enum.reduce(strings, {%{}, [], 0}, fn str, {table, list, index} ->
        if Map.has_key?(table, str) do
          {table, list, index}
        else
          {Map.put(table, str, index), list ++ [str], index + 1}
        end
      end)

    {table, unique_strings}
  end

  defp build_string_item(string) do
    escaped = escape_xml(string)

    # Preserve whitespace if the string has leading/trailing spaces
    if String.trim(string) != string do
      "<si><t xml:space=\"preserve\">#{escaped}</t></si>"
    else
      "<si><t>#{escaped}</t></si>"
    end
  end

  defp escape_xml(string) do
    string
    |> String.replace("&", "&amp;")
    |> String.replace("<", "&lt;")
    |> String.replace(">", "&gt;")
    |> String.replace("\"", "&quot;")
    |> String.replace("'", "&apos;")
  end
end
