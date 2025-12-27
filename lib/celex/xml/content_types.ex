defmodule Celex.XML.ContentTypes do
  @moduledoc """
  Generates the [Content_Types].xml file for an XLSX package.

  This file defines the MIME types for each part in the package.
  """

  @namespace "http://schemas.openxmlformats.org/package/2006/content-types"

  @content_types %{
    rels: "application/vnd.openxmlformats-package.relationships+xml",
    xml: "application/xml",
    workbook: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    worksheet: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
    shared_strings:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
    styles: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
    core_props: "application/vnd.openxmlformats-package.core-properties+xml",
    app_props: "application/vnd.openxmlformats-officedocument.extended-properties+xml"
  }

  @doc """
  Generates the [Content_Types].xml content.

  ## Options

  - `:sheet_count` - Number of worksheets (required)
  - `:include_shared_strings` - Whether to include shared strings (default: true)
  - `:include_core_props` - Whether to include core properties (default: false)
  - `:include_app_props` - Whether to include app properties (default: false)
  """
  @spec generate(keyword()) :: String.t()
  def generate(opts) do
    sheet_count = Keyword.fetch!(opts, :sheet_count)
    include_shared_strings = Keyword.get(opts, :include_shared_strings, true)
    include_core_props = Keyword.get(opts, :include_core_props, false)
    include_app_props = Keyword.get(opts, :include_app_props, false)

    defaults = build_defaults()

    overrides =
      build_overrides(sheet_count, include_shared_strings, include_core_props, include_app_props)

    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="#{@namespace}">
    #{defaults}
    #{overrides}
    </Types>
    """
    |> String.trim()
  end

  defp build_defaults do
    [
      default("rels", @content_types.rels),
      default("xml", @content_types.xml)
    ]
    |> Enum.join("\n")
  end

  defp build_overrides(sheet_count, include_shared_strings, include_core_props, include_app_props) do
    overrides = [
      override("/xl/workbook.xml", @content_types.workbook)
    ]

    # Add worksheet overrides
    sheet_overrides =
      for i <- 1..sheet_count do
        override("/xl/worksheets/sheet#{i}.xml", @content_types.worksheet)
      end

    overrides = overrides ++ sheet_overrides

    # Add optional overrides
    overrides =
      if include_shared_strings do
        overrides ++ [override("/xl/sharedStrings.xml", @content_types.shared_strings)]
      else
        overrides
      end

    overrides = overrides ++ [override("/xl/styles.xml", @content_types.styles)]

    overrides =
      if include_core_props do
        overrides ++ [override("/docProps/core.xml", @content_types.core_props)]
      else
        overrides
      end

    overrides =
      if include_app_props do
        overrides ++ [override("/docProps/app.xml", @content_types.app_props)]
      else
        overrides
      end

    Enum.join(overrides, "\n")
  end

  defp default(extension, content_type) do
    ~s(<Default Extension="#{extension}" ContentType="#{content_type}"/>)
  end

  defp override(part_name, content_type) do
    ~s(<Override PartName="#{part_name}" ContentType="#{content_type}"/>)
  end
end
