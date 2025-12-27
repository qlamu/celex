defmodule Celex.XML.WorkbookXMLTest do
  use ExUnit.Case, async: true

  alias Celex.XML.WorkbookXML

  describe "generate/1" do
    test "generates valid workbook XML" do
      xml = WorkbookXML.generate(sheet_names: ["Sheet1"])

      assert String.starts_with?(xml, "<?xml")
      assert String.contains?(xml, "<workbook")

      assert String.contains?(
               xml,
               "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
             )
    end

    test "includes sheets element" do
      xml = WorkbookXML.generate(sheet_names: ["Sheet1"])

      assert String.contains?(xml, "<sheets>")
      assert String.contains?(xml, "</sheets>")
    end

    test "includes sheet entries with names" do
      xml = WorkbookXML.generate(sheet_names: ["Data", "Summary"])

      assert String.contains?(xml, ~s(name="Data"))
      assert String.contains?(xml, ~s(name="Summary"))
    end

    test "includes sheet IDs" do
      xml = WorkbookXML.generate(sheet_names: ["Sheet1", "Sheet2"])

      assert String.contains?(xml, ~s(sheetId="1"))
      assert String.contains?(xml, ~s(sheetId="2"))
    end

    test "includes relationship IDs" do
      xml = WorkbookXML.generate(sheet_names: ["Sheet1", "Sheet2"])

      assert String.contains?(xml, ~s(r:id="rId1"))
      assert String.contains?(xml, ~s(r:id="rId2"))
    end

    test "escapes special characters in sheet names" do
      xml = WorkbookXML.generate(sheet_names: ["Data & Analysis", "Test <Sheet>"])

      assert String.contains?(xml, "Data &amp; Analysis")
      assert String.contains?(xml, "Test &lt;Sheet&gt;")
    end

    test "handles empty sheet list" do
      xml = WorkbookXML.generate(sheet_names: [])

      assert String.contains?(xml, "<sheets>")
      assert String.contains?(xml, "</sheets>")
    end
  end
end
