defmodule Celex.XML.Relationships do
  @moduledoc """
  Generates relationship XML files for an XLSX package.

  XLSX packages use relationship files to define connections between parts:
  - `_rels/.rels` - Root relationships (links to workbook)
  - `xl/_rels/workbook.xml.rels` - Workbook relationships (links to sheets, styles, strings)
  """

  @namespace "http://schemas.openxmlformats.org/package/2006/relationships"

  @relationship_types %{
    office_document:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
    styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    shared_strings:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
    core_props:
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    app_props:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
  }

  @doc """
  Generates the root _rels/.rels file.

  This file links to the main workbook document.
  """
  @spec generate_root(keyword()) :: String.t()
  def generate_root(opts \\ []) do
    include_core_props = Keyword.get(opts, :include_core_props, false)
    include_app_props = Keyword.get(opts, :include_app_props, false)

    rels = [
      relationship("rId1", @relationship_types.office_document, "xl/workbook.xml")
    ]

    rels =
      if include_core_props do
        rels ++ [relationship("rId2", @relationship_types.core_props, "docProps/core.xml")]
      else
        rels
      end

    rels =
      if include_app_props do
        next_id = if include_core_props, do: "rId3", else: "rId2"
        rels ++ [relationship(next_id, @relationship_types.app_props, "docProps/app.xml")]
      else
        rels
      end

    build_relationships_xml(rels)
  end

  @doc """
  Generates the xl/_rels/workbook.xml.rels file.

  This file links the workbook to its worksheets, styles, and shared strings.

  ## Options

  - `:sheet_count` - Number of worksheets (required)
  - `:include_shared_strings` - Whether to include shared strings (default: true)
  """
  @spec generate_workbook(keyword()) :: String.t()
  def generate_workbook(opts) do
    sheet_count = Keyword.fetch!(opts, :sheet_count)
    include_shared_strings = Keyword.get(opts, :include_shared_strings, true)

    # Worksheets get rId1, rId2, etc.
    sheet_rels =
      for i <- 1..sheet_count do
        relationship("rId#{i}", @relationship_types.worksheet, "worksheets/sheet#{i}.xml")
      end

    # Styles gets next ID
    styles_id = sheet_count + 1

    rels =
      sheet_rels ++ [relationship("rId#{styles_id}", @relationship_types.styles, "styles.xml")]

    # Shared strings gets next ID
    rels =
      if include_shared_strings do
        ss_id = styles_id + 1

        rels ++
          [relationship("rId#{ss_id}", @relationship_types.shared_strings, "sharedStrings.xml")]
      else
        rels
      end

    build_relationships_xml(rels)
  end

  defp build_relationships_xml(rels) do
    rels_content = Enum.join(rels, "\n")

    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="#{@namespace}">
    #{rels_content}
    </Relationships>
    """
    |> String.trim()
  end

  defp relationship(id, type, target) do
    ~s(<Relationship Id="#{id}" Type="#{type}" Target="#{target}"/>)
  end
end
