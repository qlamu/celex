defmodule Celex.XML.WorkbookXML do
  @moduledoc """
  Generates the xl/workbook.xml file for an XLSX package.

  This file defines the workbook structure including the list of sheets.
  """

  @main_ns "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  @rel_ns "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  @doc """
  Generates the workbook.xml content.

  ## Options

  - `:sheet_names` - List of sheet names (required)
  - `:active_sheet` - Index of active sheet (default: 0)
  """
  @spec generate(keyword()) :: String.t()
  def generate(opts) do
    sheet_names = Keyword.fetch!(opts, :sheet_names)
    active_sheet = Keyword.get(opts, :active_sheet, 0)

    sheets_xml = build_sheets(sheet_names)
    book_views = build_book_views(active_sheet)

    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="#{@main_ns}" xmlns:r="#{@rel_ns}">
    #{book_views}
    <sheets>
    #{sheets_xml}
    </sheets>
    </workbook>
    """
    |> String.trim()
  end

  defp build_sheets(sheet_names) do
    sheet_names
    |> Enum.with_index(1)
    |> Enum.map(fn {name, index} ->
      escaped_name = escape_xml(name)
      ~s(<sheet name="#{escaped_name}" sheetId="#{index}" r:id="rId#{index}"/>)
    end)
    |> Enum.join("\n")
  end

  defp build_book_views(active_sheet) do
    """
    <bookViews>
    <workbookView activeTab="#{active_sheet}"/>
    </bookViews>
    """
    |> String.trim()
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
