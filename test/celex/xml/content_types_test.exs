defmodule Celex.XML.ContentTypesTest do
  use ExUnit.Case, async: true

  alias Celex.XML.ContentTypes

  describe "generate/1" do
    test "generates valid content types XML" do
      xml = ContentTypes.generate(sheet_count: 1)

      assert String.starts_with?(xml, "<?xml")
      assert String.contains?(xml, "<Types")

      assert String.contains?(
               xml,
               "xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\""
             )
    end

    test "includes default extensions" do
      xml = ContentTypes.generate(sheet_count: 1)

      assert String.contains?(xml, "Extension=\"rels\"")
      assert String.contains?(xml, "Extension=\"xml\"")
    end

    test "includes workbook override" do
      xml = ContentTypes.generate(sheet_count: 1)

      assert String.contains?(xml, "PartName=\"/xl/workbook.xml\"")
      assert String.contains?(xml, "spreadsheetml.sheet.main+xml")
    end

    test "includes worksheet overrides for each sheet" do
      xml = ContentTypes.generate(sheet_count: 3)

      assert String.contains?(xml, "PartName=\"/xl/worksheets/sheet1.xml\"")
      assert String.contains?(xml, "PartName=\"/xl/worksheets/sheet2.xml\"")
      assert String.contains?(xml, "PartName=\"/xl/worksheets/sheet3.xml\"")
    end

    test "includes shared strings override" do
      xml = ContentTypes.generate(sheet_count: 1)

      assert String.contains?(xml, "PartName=\"/xl/sharedStrings.xml\"")
      assert String.contains?(xml, "sharedStrings+xml")
    end

    test "includes styles override" do
      xml = ContentTypes.generate(sheet_count: 1)

      assert String.contains?(xml, "PartName=\"/xl/styles.xml\"")
      assert String.contains?(xml, "styles+xml")
    end

    test "optionally excludes shared strings" do
      xml = ContentTypes.generate(sheet_count: 1, include_shared_strings: false)

      refute String.contains?(xml, "sharedStrings")
    end
  end
end
