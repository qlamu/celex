defmodule Celex.WorkbookTest do
  use ExUnit.Case, async: true

  alias Celex.Workbook
  alias Celex.Sheet

  describe "new/0" do
    test "creates empty workbook" do
      workbook = Workbook.new()
      assert workbook.sheets == []
      assert workbook.active_sheet == nil
    end
  end

  describe "add_sheet/2" do
    test "adds sheet by name" do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Sheet1")

      assert length(workbook.sheets) == 1
      assert hd(workbook.sheets).name == "Sheet1"
    end

    test "adds Sheet struct" do
      sheet = Sheet.new("Custom")

      workbook =
        Workbook.new()
        |> Workbook.add_sheet(sheet)

      assert length(workbook.sheets) == 1
      assert hd(workbook.sheets).name == "Custom"
    end

    test "adds multiple sheets" do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Sheet1")
        |> Workbook.add_sheet("Sheet2")
        |> Workbook.add_sheet("Sheet3")

      assert length(workbook.sheets) == 3
    end

    test "sets first sheet as active by default" do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Sheet1")

      assert workbook.active_sheet == "Sheet1"
    end
  end

  describe "get_sheet/2" do
    test "gets sheet by name" do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Data")

      sheet = Workbook.get_sheet(workbook, "Data")
      assert sheet.name == "Data"
    end

    test "gets sheet by index" do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Sheet1")
        |> Workbook.add_sheet("Sheet2")

      sheet = Workbook.get_sheet(workbook, 1)
      assert sheet.name == "Sheet2"
    end

    test "returns nil for non-existent sheet" do
      workbook = Workbook.new()
      assert Workbook.get_sheet(workbook, "Missing") == nil
    end
  end

  describe "update_sheet/3" do
    test "updates sheet by name" do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Data")

      updated =
        Workbook.update_sheet(workbook, "Data", fn sheet ->
          Sheet.put_cell(sheet, "A1", "Hello")
        end)

      sheet = Workbook.get_sheet(updated, "Data")
      assert Sheet.get_cell(sheet, "A1").value == "Hello"
    end
  end

  describe "sheet_names/1" do
    test "returns list of sheet names" do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Sheet1")
        |> Workbook.add_sheet("Sheet2")
        |> Workbook.add_sheet("Sheet3")

      assert Workbook.sheet_names(workbook) == ["Sheet1", "Sheet2", "Sheet3"]
    end

    test "returns empty list for empty workbook" do
      assert Workbook.sheet_names(Workbook.new()) == []
    end
  end

  describe "sheet_count/1" do
    test "returns number of sheets" do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Sheet1")
        |> Workbook.add_sheet("Sheet2")

      assert Workbook.sheet_count(workbook) == 2
    end
  end

  describe "set_active_sheet/2" do
    test "sets active sheet by name" do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Sheet1")
        |> Workbook.add_sheet("Sheet2")
        |> Workbook.set_active_sheet("Sheet2")

      assert workbook.active_sheet == "Sheet2"
    end
  end

  describe "remove_sheet/2" do
    test "removes sheet by name" do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Sheet1")
        |> Workbook.add_sheet("Sheet2")
        |> Workbook.remove_sheet("Sheet1")

      assert Workbook.sheet_count(workbook) == 1
      assert Workbook.get_sheet(workbook, "Sheet1") == nil
    end
  end

  describe "rename_sheet/3" do
    test "renames sheet" do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("OldName")
        |> Workbook.rename_sheet("OldName", "NewName")

      assert Workbook.get_sheet(workbook, "OldName") == nil
      assert Workbook.get_sheet(workbook, "NewName") != nil
    end

    test "updates active sheet if renamed" do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("OldName")
        |> Workbook.rename_sheet("OldName", "NewName")

      assert workbook.active_sheet == "NewName"
    end
  end

  describe "metadata" do
    test "sets and gets metadata" do
      workbook =
        Workbook.new()
        |> Workbook.set_metadata(:creator, "John Doe")
        |> Workbook.set_metadata(:title, "Report")

      assert Workbook.get_metadata(workbook, :creator) == "John Doe"
      assert Workbook.get_metadata(workbook, :title) == "Report"
    end
  end
end
