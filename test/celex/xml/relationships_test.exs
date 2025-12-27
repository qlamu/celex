defmodule Celex.XML.RelationshipsTest do
  use ExUnit.Case, async: true

  alias Celex.XML.Relationships

  describe "generate_root/0" do
    test "generates root .rels file" do
      xml = Relationships.generate_root()

      assert String.starts_with?(xml, "<?xml")
      assert String.contains?(xml, "<Relationships")

      assert String.contains?(
               xml,
               "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\""
             )
    end

    test "includes workbook relationship" do
      xml = Relationships.generate_root()

      assert String.contains?(xml, "xl/workbook.xml")
      assert String.contains?(xml, "officeDocument")
    end
  end

  describe "generate_workbook/1" do
    test "generates workbook.xml.rels" do
      xml = Relationships.generate_workbook(sheet_count: 1)

      assert String.starts_with?(xml, "<?xml")
      assert String.contains?(xml, "<Relationships")
    end

    test "includes worksheet relationships" do
      xml = Relationships.generate_workbook(sheet_count: 3)

      assert String.contains?(xml, "worksheets/sheet1.xml")
      assert String.contains?(xml, "worksheets/sheet2.xml")
      assert String.contains?(xml, "worksheets/sheet3.xml")
      assert String.contains?(xml, "worksheet")
    end

    test "includes styles relationship" do
      xml = Relationships.generate_workbook(sheet_count: 1)

      assert String.contains?(xml, "styles.xml")
    end

    test "includes shared strings relationship by default" do
      xml = Relationships.generate_workbook(sheet_count: 1)

      assert String.contains?(xml, "sharedStrings.xml")
    end

    test "optionally excludes shared strings" do
      xml = Relationships.generate_workbook(sheet_count: 1, include_shared_strings: false)

      refute String.contains?(xml, "sharedStrings")
    end

    test "has unique rId values" do
      xml = Relationships.generate_workbook(sheet_count: 3)

      # Should have rId1, rId2, rId3 for sheets, plus styles and shared strings
      assert String.contains?(xml, "rId1")
      assert String.contains?(xml, "rId2")
      assert String.contains?(xml, "rId3")
      assert String.contains?(xml, "rId4")
      assert String.contains?(xml, "rId5")
    end
  end
end
