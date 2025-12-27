defmodule Celex.SheetTest do
  use ExUnit.Case, async: true

  alias Celex.Sheet
  alias Celex.Row

  describe "new/1" do
    test "creates sheet with name" do
      sheet = Sheet.new("Data")
      assert sheet.name == "Data"
      assert sheet.cells == %{}
      assert sheet.merged_regions == []
    end
  end

  describe "new/2 with options" do
    test "creates sheet with options" do
      sheet = Sheet.new("Data", index: 1, hidden: true)
      assert sheet.name == "Data"
      assert sheet.index == 1
      assert sheet.hidden == true
    end
  end

  describe "put_cell/4" do
    test "puts cell at A1 notation" do
      sheet = Sheet.new("Data")
      sheet = Sheet.put_cell(sheet, "A1", "Hello")

      assert Sheet.get_cell(sheet, "A1").value == "Hello"
    end

    test "puts cell at row/col tuple" do
      sheet = Sheet.new("Data")
      sheet = Sheet.put_cell(sheet, {1, 1}, "Hello")

      assert Sheet.get_cell(sheet, {1, 1}).value == "Hello"
    end

    test "puts cell with style" do
      sheet = Sheet.new("Data")
      sheet = Sheet.put_cell(sheet, "A1", "Bold", style: %{bold: true})

      cell = Sheet.get_cell(sheet, "A1")
      assert cell.value == "Bold"
      assert cell.style == %{bold: true}
    end

    test "puts cell with formula" do
      sheet = Sheet.new("Data")
      sheet = Sheet.put_cell(sheet, "A3", nil, formula: "SUM(A1:A2)")

      cell = Sheet.get_cell(sheet, "A3")
      assert cell.formula == "SUM(A1:A2)"
      assert cell.type == :formula
    end

    test "overwrites existing cell" do
      sheet =
        Sheet.new("Data")
        |> Sheet.put_cell("A1", "First")
        |> Sheet.put_cell("A1", "Second")

      assert Sheet.get_cell(sheet, "A1").value == "Second"
    end
  end

  describe "get_cell/2" do
    test "gets cell by A1 notation" do
      sheet =
        Sheet.new("Data")
        |> Sheet.put_cell("B2", 42)

      cell = Sheet.get_cell(sheet, "B2")
      assert cell.value == 42
    end

    test "returns nil for empty cell" do
      sheet = Sheet.new("Data")
      assert Sheet.get_cell(sheet, "Z99") == nil
    end
  end

  describe "put_row/3" do
    test "puts row at index" do
      sheet = Sheet.new("Data")
      sheet = Sheet.put_row(sheet, 1, ["A", "B", "C"])

      assert Sheet.get_cell(sheet, "A1").value == "A"
      assert Sheet.get_cell(sheet, "B1").value == "B"
      assert Sheet.get_cell(sheet, "C1").value == "C"
    end

    test "puts Row struct" do
      sheet = Sheet.new("Data")
      row = Row.new(["X", "Y", "Z"])
      sheet = Sheet.put_row(sheet, 2, row)

      assert Sheet.get_cell(sheet, "A2").value == "X"
      assert Sheet.get_cell(sheet, "B2").value == "Y"
    end
  end

  describe "get_row/2" do
    test "gets row as list of cells" do
      sheet =
        Sheet.new("Data")
        |> Sheet.put_cell("A1", "A")
        |> Sheet.put_cell("B1", "B")
        |> Sheet.put_cell("C1", "C")

      row = Sheet.get_row(sheet, 1)
      assert length(row) == 3
      assert Enum.at(row, 0).value == "A"
    end

    test "returns empty list for empty row" do
      sheet = Sheet.new("Data")
      assert Sheet.get_row(sheet, 99) == []
    end
  end

  describe "set_column_width/3" do
    test "sets width for column by letter" do
      sheet =
        Sheet.new("Data")
        |> Sheet.set_column_width("A", 20)

      assert sheet.column_widths["A"] == 20
    end

    test "sets width for column by index" do
      sheet =
        Sheet.new("Data")
        |> Sheet.set_column_width(1, 15.5)

      assert sheet.column_widths[1] == 15.5
    end
  end

  describe "set_row_height/3" do
    test "sets height for row" do
      sheet =
        Sheet.new("Data")
        |> Sheet.set_row_height(1, 25)

      assert sheet.row_heights[1] == 25
    end
  end

  describe "merge_cells/2" do
    test "adds merged region" do
      sheet =
        Sheet.new("Data")
        |> Sheet.merge_cells("A1:B2")

      assert "A1:B2" in sheet.merged_regions
    end

    test "adds multiple merged regions" do
      sheet =
        Sheet.new("Data")
        |> Sheet.merge_cells("A1:B2")
        |> Sheet.merge_cells("D1:E2")

      assert length(sheet.merged_regions) == 2
    end
  end

  describe "cell_count/1" do
    test "returns number of cells" do
      sheet =
        Sheet.new("Data")
        |> Sheet.put_cell("A1", "A")
        |> Sheet.put_cell("B2", "B")
        |> Sheet.put_cell("C3", "C")

      assert Sheet.cell_count(sheet) == 3
    end

    test "returns 0 for empty sheet" do
      sheet = Sheet.new("Data")
      assert Sheet.cell_count(sheet) == 0
    end
  end

  describe "dimensions/1" do
    test "returns dimensions of used area" do
      sheet =
        Sheet.new("Data")
        |> Sheet.put_cell("A1", "A")
        |> Sheet.put_cell("C5", "C")

      {start_cell, end_cell} = Sheet.dimensions(sheet)
      assert start_cell == {1, 1}
      assert end_cell == {5, 3}
    end

    test "returns nil for empty sheet" do
      sheet = Sheet.new("Data")
      assert Sheet.dimensions(sheet) == nil
    end
  end

  describe "rows/1" do
    test "returns stream of rows" do
      sheet =
        Sheet.new("Data")
        |> Sheet.put_cell("A1", "A1")
        |> Sheet.put_cell("B1", "B1")
        |> Sheet.put_cell("A2", "A2")
        |> Sheet.put_cell("B2", "B2")

      rows = Sheet.rows(sheet) |> Enum.to_list()
      assert length(rows) == 2

      [row1, row2] = rows
      assert Enum.at(row1, 0).value == "A1"
      assert Enum.at(row1, 1).value == "B1"
      assert Enum.at(row2, 0).value == "A2"
    end
  end
end
