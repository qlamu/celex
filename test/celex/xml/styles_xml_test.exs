defmodule Celex.XML.StylesXMLTest do
  use ExUnit.Case, async: true

  alias Celex.XML.StylesXML

  describe "generate/0" do
    test "generates valid styles XML with defaults" do
      xml = StylesXML.generate()

      assert String.starts_with?(xml, "<?xml")
      assert String.contains?(xml, "<styleSheet")

      assert String.contains?(
               xml,
               "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
             )
    end

    test "includes numFmts section" do
      xml = StylesXML.generate()

      assert String.contains?(xml, "<numFmts")
    end

    test "includes fonts section with default font" do
      xml = StylesXML.generate()

      assert String.contains?(xml, "<fonts")
      assert String.contains?(xml, "<font>")
    end

    test "includes fills section with default fills" do
      xml = StylesXML.generate()

      assert String.contains?(xml, "<fills")
      assert String.contains?(xml, "patternFill")
    end

    test "includes borders section with default border" do
      xml = StylesXML.generate()

      assert String.contains?(xml, "<borders")
      assert String.contains?(xml, "<border>")
    end

    test "includes cellXfs section" do
      xml = StylesXML.generate()

      assert String.contains?(xml, "<cellXfs")
      assert String.contains?(xml, "<xf")
    end
  end

  describe "generate/1 with styles" do
    test "includes bold font" do
      styles = [%{font: %{bold: true}}]
      xml = StylesXML.generate(styles)

      assert String.contains?(xml, "<b/>")
    end

    test "includes italic font" do
      styles = [%{font: %{italic: true}}]
      xml = StylesXML.generate(styles)

      assert String.contains?(xml, "<i/>")
    end

    test "includes font size" do
      styles = [%{font: %{size: 14}}]
      xml = StylesXML.generate(styles)

      assert String.contains?(xml, ~s(sz val="14"))
    end

    test "includes font color" do
      styles = [%{font: %{color: "#FF0000"}}]
      xml = StylesXML.generate(styles)

      assert String.contains?(xml, ~s(rgb="FFFF0000"))
    end

    test "includes fill color" do
      styles = [%{fill: %{color: "#FFFF00", pattern: :solid}}]
      xml = StylesXML.generate(styles)

      assert String.contains?(xml, "patternType=\"solid\"")
      assert String.contains?(xml, "FFFFFF00")
    end

    test "includes border styles" do
      styles = [%{border: %{left: :thin, right: :thin, top: :thin, bottom: :thin}}]
      xml = StylesXML.generate(styles)

      assert String.contains?(xml, "<left style=\"thin\"")
      assert String.contains?(xml, "<right style=\"thin\"")
      assert String.contains?(xml, "<top style=\"thin\"")
      assert String.contains?(xml, "<bottom style=\"thin\"")
    end

    test "includes border color" do
      styles = [%{border: %{left: :thin, color: "#000000"}}]
      xml = StylesXML.generate(styles)

      assert String.contains?(xml, "style=\"thin\"")
      assert String.contains?(xml, "FF000000")
    end

    test "includes medium and thick borders" do
      styles = [%{border: %{left: :medium, right: :thick}}]
      xml = StylesXML.generate(styles)

      assert String.contains?(xml, "style=\"medium\"")
      assert String.contains?(xml, "style=\"thick\"")
    end

    test "includes custom number format" do
      styles = [%{number_format: "0.000"}]
      xml = StylesXML.generate(styles)

      assert String.contains?(xml, "formatCode=\"0.000\"")
      assert String.contains?(xml, "numFmtId=\"164\"")
    end

    test "includes multiple custom number formats" do
      styles = [
        %{number_format: "0.00%"},
        %{number_format: "#,##0.00"}
      ]

      xml = StylesXML.generate(styles)

      assert String.contains?(xml, "formatCode=\"0.00%\"")
      assert String.contains?(xml, "formatCode=\"#,##0.00\"")
      assert String.contains?(xml, "numFmtId=\"164\"")
      assert String.contains?(xml, "numFmtId=\"165\"")
    end
  end

  describe "build_style_index/1" do
    test "builds style to index mapping" do
      styles = [
        %{font: %{bold: true}},
        %{fill: %{color: "#FF0000"}},
        # Duplicate
        %{font: %{bold: true}}
      ]

      {index_map, unique_count} = StylesXML.build_style_index(styles)

      # Duplicates should have same index
      assert unique_count == 2
      assert is_map(index_map)
    end
  end
end
