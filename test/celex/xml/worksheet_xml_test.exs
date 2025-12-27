defmodule Celex.XML.WorksheetXMLTest do
  use ExUnit.Case, async: true

  alias Celex.XML.WorksheetXML
  alias Celex.Sheet
  alias Celex.Cell

  describe "generate/1" do
    test "generates valid worksheet XML" do
      sheet = Sheet.new("Test")
      xml = WorksheetXML.generate(sheet)

      assert String.starts_with?(xml, "<?xml")
      assert String.contains?(xml, "<worksheet")

      assert String.contains?(
               xml,
               "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
             )
    end

    test "includes sheetData element" do
      sheet = Sheet.new("Test")
      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "<sheetData")
    end

    test "generates row and cell for string value" do
      sheet =
        Sheet.new("Test")
        |> Sheet.put_cell("A1", "Hello")

      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "<row r=\"1\"")
      assert String.contains?(xml, "<c r=\"A1\"")
      assert String.contains?(xml, "t=\"inlineStr\"")
      assert String.contains?(xml, "<is><t>Hello</t></is>")
    end

    test "generates cell for number value" do
      sheet =
        Sheet.new("Test")
        |> Sheet.put_cell("A1", 42)

      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "<c r=\"A1\"")
      assert String.contains?(xml, "<v>42</v>")
      refute String.contains?(xml, "t=\"inlineStr\"")
    end

    test "generates cell for float value" do
      sheet =
        Sheet.new("Test")
        |> Sheet.put_cell("A1", 3.14159)

      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "<v>3.14159</v>")
    end

    test "generates cell for boolean true" do
      sheet =
        Sheet.new("Test")
        |> Sheet.put_cell("A1", true)

      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "t=\"b\"")
      assert String.contains?(xml, "<v>1</v>")
    end

    test "generates cell for boolean false" do
      sheet =
        Sheet.new("Test")
        |> Sheet.put_cell("A1", false)

      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "t=\"b\"")
      assert String.contains?(xml, "<v>0</v>")
    end

    test "generates cell for date" do
      sheet =
        Sheet.new("Test")
        |> Sheet.put_cell("A1", ~D[2024-06-15])

      xml = WorksheetXML.generate(sheet)

      # Date should be converted to Excel serial number
      assert String.contains?(xml, "<v>")
    end

    test "generates cell for formula" do
      sheet =
        Sheet.new("Test")
        |> Sheet.put_cell("A3", nil, formula: "SUM(A1:A2)")

      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "<f>SUM(A1:A2)</f>")
    end

    test "handles multiple cells in a row" do
      sheet =
        Sheet.new("Test")
        |> Sheet.put_cell("A1", "A")
        |> Sheet.put_cell("B1", "B")
        |> Sheet.put_cell("C1", "C")

      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "<c r=\"A1\"")
      assert String.contains?(xml, "<c r=\"B1\"")
      assert String.contains?(xml, "<c r=\"C1\"")
    end

    test "handles multiple rows" do
      sheet =
        Sheet.new("Test")
        |> Sheet.put_cell("A1", "Row 1")
        |> Sheet.put_cell("A2", "Row 2")
        |> Sheet.put_cell("A3", "Row 3")

      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "<row r=\"1\"")
      assert String.contains?(xml, "<row r=\"2\"")
      assert String.contains?(xml, "<row r=\"3\"")
    end

    test "escapes special XML characters in strings" do
      sheet =
        Sheet.new("Test")
        |> Sheet.put_cell("A1", "Data & More <test>")

      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "Data &amp; More &lt;test&gt;")
    end

    test "handles empty sheet" do
      sheet = Sheet.new("Empty")
      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "<sheetData")
      assert String.contains?(xml, "</sheetData>")
    end

    test "includes dimension if sheet has cells" do
      sheet =
        Sheet.new("Test")
        |> Sheet.put_cell("A1", "Start")
        |> Sheet.put_cell("C5", "End")

      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "<dimension ref=\"A1:C5\"")
    end

    test "includes column widths" do
      sheet =
        Sheet.new("Test")
        |> Sheet.set_column_width("A", 20)
        |> Sheet.set_column_width("B", 30)
        |> Sheet.put_cell("A1", "Test")

      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "<cols>")
      assert String.contains?(xml, "width=\"20\"")
      assert String.contains?(xml, "width=\"30\"")
    end

    test "includes merged cells" do
      sheet =
        Sheet.new("Test")
        |> Sheet.put_cell("A1", "Merged")
        |> Sheet.merge_cells("A1:B2")

      xml = WorksheetXML.generate(sheet)

      assert String.contains?(xml, "<mergeCells")
      assert String.contains?(xml, "<mergeCell ref=\"A1:B2\"")
    end
  end

  describe "generate_row/2" do
    test "generates row XML from values" do
      xml = WorksheetXML.generate_row(1, ["A", "B", "C"])

      assert String.contains?(xml, "<row r=\"1\"")
      assert String.contains?(xml, "<c r=\"A1\"")
      assert String.contains?(xml, "<c r=\"B1\"")
      assert String.contains?(xml, "<c r=\"C1\"")
    end

    test "generates row XML from Cell structs" do
      cells = [
        Cell.new("Hello"),
        Cell.new(42),
        Cell.new(true)
      ]

      xml = WorksheetXML.generate_row(1, cells)

      assert String.contains?(xml, "<is><t>Hello</t></is>")
      assert String.contains?(xml, "<v>42</v>")
      assert String.contains?(xml, "t=\"b\"")
    end
  end
end
