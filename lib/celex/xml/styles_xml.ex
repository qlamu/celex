defmodule Celex.XML.StylesXML do
  @moduledoc """
  Generates the xl/styles.xml file for an XLSX package.

  Excel styles are complex and indexed. This module provides:
  - Default styles for basic functionality
  - Custom styles for fonts, fills, borders, and number formats
  """

  @main_ns "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  # Default number formats built into Excel
  @builtin_num_formats %{
    general: 0,
    integer: 1,
    decimal: 2,
    percentage: 9,
    date: 14,
    time: 21,
    datetime: 22
  }

  @doc """
  Generates the styles.xml content.

  With no arguments, generates minimal default styles.
  With a list of style maps, generates styles for those configurations.
  """
  @spec generate([map()]) :: String.t()
  def generate(styles \\ []) do
    {fonts_xml, font_count} = build_fonts(styles)
    {fills_xml, fill_count} = build_fills(styles)
    {borders_xml, border_count} = build_borders(styles)
    {num_fmts_xml, num_fmt_count} = build_number_formats(styles)
    {cell_xfs_xml, xf_count} = build_cell_xfs(styles, font_count, fill_count)

    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="#{@main_ns}">
    <numFmts count="#{num_fmt_count}">
    #{num_fmts_xml}
    </numFmts>
    <fonts count="#{font_count}">
    #{fonts_xml}
    </fonts>
    <fills count="#{fill_count}">
    #{fills_xml}
    </fills>
    <borders count="#{border_count}">
    #{borders_xml}
    </borders>
    <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="#{xf_count}">
    #{cell_xfs_xml}
    </cellXfs>
    </styleSheet>
    """
    |> String.trim()
  end

  @doc """
  Builds a style index mapping from a list of styles.

  Returns {index_map, unique_count} where index_map maps each unique
  style to its index in the cellXfs section.
  """
  @spec build_style_index([map()]) :: {map(), non_neg_integer()}
  def build_style_index(styles) do
    {unique_styles, _} =
      Enum.reduce(styles, {[], MapSet.new()}, fn style, {list, seen} ->
        key = style_key(style)

        if MapSet.member?(seen, key) do
          {list, seen}
        else
          {list ++ [style], MapSet.put(seen, key)}
        end
      end)

    # Index 0 is the default style, so custom styles start at 1
    index_map =
      unique_styles
      |> Enum.with_index(1)
      |> Enum.into(%{}, fn {style, idx} -> {style_key(style), idx} end)

    {index_map, length(unique_styles)}
  end

  defp style_key(style), do: :erlang.phash2(style)

  # Fonts

  defp build_fonts(styles) do
    # Default font (required)
    default_font = "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>"

    custom_fonts =
      styles
      |> Enum.filter(&Map.has_key?(&1, :font))
      |> Enum.map(&build_font(&1.font))
      |> Enum.uniq()

    all_fonts = [default_font | custom_fonts]
    {Enum.join(all_fonts, "\n"), length(all_fonts)}
  end

  defp build_font(font) when is_map(font) do
    elements = []

    elements = if font[:bold], do: elements ++ ["<b/>"], else: elements
    elements = if font[:italic], do: elements ++ ["<i/>"], else: elements
    elements = if font[:underline], do: elements ++ ["<u/>"], else: elements
    elements = if font[:strikethrough], do: elements ++ ["<strike/>"], else: elements

    elements =
      if font[:size] do
        elements ++ [~s(<sz val="#{font[:size]}"/>)]
      else
        elements ++ [~s(<sz val="11"/>)]
      end

    elements =
      if font[:color] do
        rgb = normalize_color(font[:color])
        elements ++ [~s(<color rgb="#{rgb}"/>)]
      else
        elements
      end

    elements =
      if font[:name] do
        elements ++ [~s(<name val="#{font[:name]}"/>)]
      else
        elements ++ [~s(<name val="Calibri"/>)]
      end

    "<font>#{Enum.join(elements)}</font>"
  end

  # Fills

  defp build_fills(styles) do
    # Excel requires two default fills
    default_fills = [
      "<fill><patternFill patternType=\"none\"/></fill>",
      "<fill><patternFill patternType=\"gray125\"/></fill>"
    ]

    custom_fills =
      styles
      |> Enum.filter(&Map.has_key?(&1, :fill))
      |> Enum.map(&build_fill(&1.fill))
      |> Enum.uniq()

    all_fills = default_fills ++ custom_fills
    {Enum.join(all_fills, "\n"), length(all_fills)}
  end

  defp build_fill(fill) when is_map(fill) do
    pattern = fill[:pattern] || :solid
    color = fill[:color]

    if color do
      rgb = normalize_color(color)

      """
      <fill><patternFill patternType="#{pattern}"><fgColor rgb="#{rgb}"/></patternFill></fill>
      """
      |> String.trim()
    else
      "<fill><patternFill patternType=\"#{pattern}\"/></fill>"
    end
  end

  # Borders

  defp build_borders(styles) do
    # Default empty border (required)
    default_border = "<border><left/><right/><top/><bottom/><diagonal/></border>"

    custom_borders =
      styles
      |> Enum.filter(&Map.has_key?(&1, :border))
      |> Enum.map(&build_border(&1.border))
      |> Enum.uniq()

    all_borders = [default_border | custom_borders]
    {Enum.join(all_borders, "\n"), length(all_borders)}
  end

  defp build_border(border) when is_map(border) do
    color = border[:color]
    color_xml = if color, do: "<color rgb=\"#{normalize_color(color)}\"/>", else: ""

    left = build_side_element("left", border[:left], color_xml)
    right = build_side_element("right", border[:right], color_xml)
    top = build_side_element("top", border[:top], color_xml)
    bottom = build_side_element("bottom", border[:bottom], color_xml)
    diagonal = build_side_element("diagonal", border[:diagonal], color_xml)

    "<border>#{left}#{right}#{top}#{bottom}#{diagonal}</border>"
  end

  defp build_side_element(name, nil, _color_xml), do: "<#{name}/>"
  defp build_side_element(name, :none, _color_xml), do: "<#{name}/>"

  defp build_side_element(name, style, color_xml) when is_atom(style) do
    if color_xml == "" do
      "<#{name} style=\"#{style}\"/>"
    else
      "<#{name} style=\"#{style}\">#{color_xml}</#{name}>"
    end
  end

  # Number formats

  defp build_number_formats(styles) do
    custom_formats =
      styles
      |> Enum.filter(&Map.has_key?(&1, :number_format))
      |> Enum.map(& &1.number_format)
      |> Enum.uniq()
      |> Enum.reject(&is_builtin_format?/1)

    if custom_formats == [] do
      {"", 0}
    else
      fmts =
        custom_formats
        # Custom formats start at 164
        |> Enum.with_index(164)
        |> Enum.map(fn {fmt, id} ->
          ~s(<numFmt numFmtId="#{id}" formatCode="#{escape_xml(fmt)}"/>)
        end)
        |> Enum.join("\n")

      {fmts, length(custom_formats)}
    end
  end

  defp is_builtin_format?(format) when is_atom(format),
    do: Map.has_key?(@builtin_num_formats, format)

  defp is_builtin_format?(_), do: false

  # Cell XFs (the actual style combinations)

  defp build_cell_xfs(styles, _font_count, _fill_count) do
    # Default style
    default_xf = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"

    if styles == [] do
      {default_xf, 1}
    else
      # Build XF entries for each unique style
      custom_xfs =
        styles
        |> Enum.uniq_by(&style_key/1)
        |> Enum.with_index(1)
        |> Enum.map(fn {style, _idx} ->
          build_xf(style)
        end)

      all_xfs = [default_xf | custom_xfs]
      {Enum.join(all_xfs, "\n"), length(all_xfs)}
    end
  end

  defp build_xf(style) do
    # For simplicity, we use fixed IDs for custom styles:
    # fontId=1 for custom fonts, fillId=2 for custom fills, borderId=1 for custom borders
    font_id = if style[:font], do: "1", else: "0"
    fill_id = if style[:fill], do: "2", else: "0"
    border_id = if style[:border], do: "1", else: "0"
    num_fmt_id = get_num_fmt_id(style[:number_format])

    apply_font = if style[:font], do: ~s( applyFont="1"), else: ""
    apply_fill = if style[:fill], do: ~s( applyFill="1"), else: ""
    apply_border = if style[:border], do: ~s( applyBorder="1"), else: ""
    apply_num_fmt = if style[:number_format], do: ~s( applyNumberFormat="1"), else: ""

    "<xf numFmtId=\"#{num_fmt_id}\" fontId=\"#{font_id}\" fillId=\"#{fill_id}\" borderId=\"#{border_id}\"#{apply_font}#{apply_fill}#{apply_border}#{apply_num_fmt}/>"
  end

  defp get_num_fmt_id(nil), do: 0
  defp get_num_fmt_id(format) when is_atom(format), do: Map.get(@builtin_num_formats, format, 0)
  # First custom format ID
  defp get_num_fmt_id(_custom), do: 164

  defp normalize_color("#" <> hex), do: "FF" <> String.upcase(hex)
  defp normalize_color(color), do: "FF" <> String.upcase(color)

  defp escape_xml(string) do
    string
    |> String.replace("&", "&amp;")
    |> String.replace("<", "&lt;")
    |> String.replace(">", "&gt;")
    |> String.replace("\"", "&quot;")
  end
end
