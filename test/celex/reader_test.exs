defmodule Celex.ReaderTest do
  use ExUnit.Case, async: true

  alias Celex.Reader
  alias Celex.Workbook
  alias Celex.Sheet

  setup do
    path = Path.join(System.tmp_dir!(), "reader_test_#{:erlang.unique_integer([:positive])}.xlsx")
    on_exit(fn -> File.rm(path) end)
    {:ok, path: path}
  end

  describe "read/1" do
    test "reads a simple XLSX file", %{path: path} do
      # Create a test file
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["Name", "Value"])
        |> Celex.StreamWriter.write_row(["Alice", 100])
      end)

      {:ok, workbook} = Reader.read(path)

      assert %Workbook{} = workbook
      assert Workbook.sheet_names(workbook) == ["Data"]
    end

    test "returns error for non-existent file" do
      result = Reader.read("/non/existent/file.xlsx")
      assert {:error, _reason} = result
    end

    test "reads multiple sheets", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("First")
        |> Celex.StreamWriter.write_row(["A"])
        |> Celex.StreamWriter.start_sheet("Second")
        |> Celex.StreamWriter.write_row(["B"])
      end)

      {:ok, workbook} = Reader.read(path)

      assert Workbook.sheet_names(workbook) == ["First", "Second"]
    end
  end

  describe "read!/1" do
    test "reads file and returns workbook", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["Test"])
      end)

      workbook = Reader.read!(path)
      assert %Workbook{} = workbook
    end

    test "raises on error" do
      assert_raise File.Error, fn ->
        Reader.read!("/non/existent/file.xlsx")
      end
    end
  end

  describe "get_sheet/2" do
    test "gets sheet by name", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Sales")
        |> Celex.StreamWriter.write_row(["Product", "Amount"])
        |> Celex.StreamWriter.write_row(["Widget", 42])
      end)

      {:ok, workbook} = Reader.read(path)
      {:ok, sheet} = Reader.get_sheet(workbook, "Sales")

      assert %Sheet{} = sheet
      assert sheet.name == "Sales"
    end

    test "gets sheet by index", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("First")
        |> Celex.StreamWriter.write_row(["A"])
        |> Celex.StreamWriter.start_sheet("Second")
        |> Celex.StreamWriter.write_row(["B"])
      end)

      {:ok, workbook} = Reader.read(path)
      {:ok, sheet} = Reader.get_sheet(workbook, 1)

      assert sheet.name == "Second"
    end

    test "returns error for unknown sheet name", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["Test"])
      end)

      {:ok, workbook} = Reader.read(path)
      result = Reader.get_sheet(workbook, "Unknown")

      assert {:error, :sheet_not_found} = result
    end

    test "returns error for out of bounds index", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["Test"])
      end)

      {:ok, workbook} = Reader.read(path)
      result = Reader.get_sheet(workbook, 99)

      assert {:error, :sheet_not_found} = result
    end
  end

  describe "get_cell/3" do
    test "gets cell value by reference", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["Hello", 42, true])
      end)

      {:ok, workbook} = Reader.read(path)
      {:ok, sheet} = Reader.get_sheet(workbook, "Data")

      assert Reader.get_cell(sheet, "A1") == "Hello"
      assert Reader.get_cell(sheet, "B1") == 42
      assert Reader.get_cell(sheet, "C1") == true
    end

    test "gets cell by row and column tuple", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["A1", "B1"])
        |> Celex.StreamWriter.write_row(["A2", "B2"])
      end)

      {:ok, workbook} = Reader.read(path)
      {:ok, sheet} = Reader.get_sheet(workbook, "Data")

      assert Reader.get_cell(sheet, {1, 1}) == "A1"
      assert Reader.get_cell(sheet, {1, 2}) == "B1"
      assert Reader.get_cell(sheet, {2, 1}) == "A2"
      assert Reader.get_cell(sheet, {2, 2}) == "B2"
    end

    test "returns nil for empty cell", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["A", nil, "C"])
      end)

      {:ok, workbook} = Reader.read(path)
      {:ok, sheet} = Reader.get_sheet(workbook, "Data")

      assert Reader.get_cell(sheet, "B1") == nil
      assert Reader.get_cell(sheet, "Z99") == nil
    end

    test "returns nil for invalid cell reference", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["Test"])
      end)

      {:ok, workbook} = Reader.read(path)
      {:ok, sheet} = Reader.get_sheet(workbook, "Data")

      assert Reader.get_cell(sheet, "invalid") == nil
      assert Reader.get_cell(sheet, "") == nil
    end
  end

  describe "get_rows/1" do
    test "returns all rows as lists", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["Name", "Age"])
        |> Celex.StreamWriter.write_row(["Alice", 30])
        |> Celex.StreamWriter.write_row(["Bob", 25])
      end)

      {:ok, workbook} = Reader.read(path)
      {:ok, sheet} = Reader.get_sheet(workbook, "Data")
      rows = Reader.get_rows(sheet)

      assert rows == [
               ["Name", "Age"],
               ["Alice", 30],
               ["Bob", 25]
             ]
    end

    test "handles sparse data", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["A", nil, "C"])
        |> Celex.StreamWriter.write_row([nil, "B2", nil])
      end)

      {:ok, workbook} = Reader.read(path)
      {:ok, sheet} = Reader.get_sheet(workbook, "Data")
      rows = Reader.get_rows(sheet)

      assert length(rows) == 2
      assert Enum.at(rows, 0) == ["A", nil, "C"]
      # Second row should have proper nil handling
      row2 = Enum.at(rows, 1)
      assert Enum.at(row2, 0) == nil
      assert Enum.at(row2, 1) == "B2"
    end
  end

  describe "data type handling" do
    test "preserves string values", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["Hello", "World", "Test 123"])
      end)

      {:ok, workbook} = Reader.read(path)
      {:ok, sheet} = Reader.get_sheet(workbook, "Data")

      assert Reader.get_cell(sheet, "A1") == "Hello"
      assert Reader.get_cell(sheet, "B1") == "World"
      assert Reader.get_cell(sheet, "C1") == "Test 123"
    end

    test "preserves numeric values", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row([42, 3.14, -100, 0])
      end)

      {:ok, workbook} = Reader.read(path)
      {:ok, sheet} = Reader.get_sheet(workbook, "Data")

      assert Reader.get_cell(sheet, "A1") == 42
      assert Reader.get_cell(sheet, "B1") == 3.14
      assert Reader.get_cell(sheet, "C1") == -100
      assert Reader.get_cell(sheet, "D1") == 0
    end

    test "preserves boolean values", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row([true, false])
      end)

      {:ok, workbook} = Reader.read(path)
      {:ok, sheet} = Reader.get_sheet(workbook, "Data")

      assert Reader.get_cell(sheet, "A1") == true
      assert Reader.get_cell(sheet, "B1") == false
    end
  end

  describe "round-trip with builder API" do
    test "data written with builder can be read back", %{path: path} do
      Celex.new()
      |> Celex.add_sheet("Sales")
      |> Celex.put_cell("A1", "Product")
      |> Celex.put_cell("B1", "Price")
      |> Celex.put_cell("A2", "Widget")
      |> Celex.put_cell("B2", 19.99)
      |> Celex.write!(path)

      {:ok, workbook} = Reader.read(path)
      {:ok, sheet} = Reader.get_sheet(workbook, "Sales")

      assert Reader.get_cell(sheet, "A1") == "Product"
      assert Reader.get_cell(sheet, "B1") == "Price"
      assert Reader.get_cell(sheet, "A2") == "Widget"
      assert Reader.get_cell(sheet, "B2") == 19.99
    end
  end
end
