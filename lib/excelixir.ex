defmodule Excelixir do
  @moduledoc """
  Creates Excel XLSX files without external dependencies.
  Uses Office Open XML format.
  """

  @content_types """
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  </Types>
  """

  @rels """
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  </Relationships>
  """

  @workbook """
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
      <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
  </workbook>
  """

  @workbook_rels """
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  </Relationships>
  """

  @doc """
  Creates an Excel file with the given rows and saves it to the specified path.
  """
  def create_excel(filename, rows) do
    # Generate all required XML content
    worksheet_xml = generate_worksheet_xml(rows)

    # Create ZIP file directly without temporary directory
    files = [
      {~c"[Content_Types].xml", @content_types},
      {~c"_rels/.rels", @rels},
      {~c"xl/workbook.xml", @workbook},
      {~c"xl/_rels/workbook.xml.rels", @workbook_rels},
      {~c"xl/worksheets/sheet1.xml", worksheet_xml}
    ]

    # Create ZIP file with all contents
    :zip.create(String.to_charlist(filename), files)
  end

  defp generate_worksheet_xml(rows) do
    rows_xml =
      rows
      |> Enum.with_index(1)
      |> Enum.map(&generate_row_xml/1)
      |> Enum.join("\n")

    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
        #{rows_xml}
      </sheetData>
    </worksheet>
    """
  end

  defp generate_row_xml({row, row_num}) do
    cells =
      row
      |> Enum.with_index(?A)
      |> Enum.map(fn {value, col} ->
        cell_ref = "#{<<col>>}#{row_num}"
        generate_cell_xml(value, cell_ref)
      end)
      |> Enum.join("\n")

    """
      <row r="#{row_num}">
        #{cells}
      </row>
    """
  end

  defp generate_cell_xml(value, cell_ref) do
    {type, value} =
      case value do
        v when is_number(v) -> {"n", "#{v}"}
        v when is_binary(v) -> {"inlineStr", "<is><t>#{escape_xml(v)}</t></is>"}
        v -> {"inlineStr", "<is><t>#{escape_xml("#{v}")}</t></is>"}
      end

    """
        <c r="#{cell_ref}" t="#{type}">
          #{if type == "n", do: "<v>#{value}</v>", else: value}
        </c>
    """
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
