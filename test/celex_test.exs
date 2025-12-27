defmodule CelexTest do
  use ExUnit.Case, async: true

  @test_output_dir System.tmp_dir!()

  setup do
    filename = "celex_api_test_#{:erlang.unique_integer([:positive])}.xlsx"
    path = Path.join(@test_output_dir, filename)

    on_exit(fn ->
      File.rm(path)
    end)

    {:ok, path: path}
  end

  describe "builder API" do
    test "creates workbook with new/0" do
      workbook = Celex.new()
      assert Celex.sheet_names(workbook) == []
    end

    test "adds sheet with add_sheet/2" do
      workbook =
        Celex.new()
        |> Celex.add_sheet("Data")

      assert Celex.sheet_names(workbook) == ["Data"]
    end

    test "puts cells with put_cell/3" do
      workbook =
        Celex.new()
        |> Celex.add_sheet("Data")
        |> Celex.put_cell("A1", "Hello")
        |> Celex.put_cell("B1", 42)

      assert Celex.get_cell(workbook, "Data", "A1").value == "Hello"
      assert Celex.get_cell(workbook, "Data", "B1").value == 42
    end

    test "puts cell with formula" do
      workbook =
        Celex.new()
        |> Celex.add_sheet("Data")
        |> Celex.put_cell("A1", 10)
        |> Celex.put_cell("A2", 20)
        |> Celex.put_cell("A3", nil, formula: "SUM(A1:A2)")

      cell = Celex.get_cell(workbook, "Data", "A3")
      assert cell.formula == "SUM(A1:A2)"
    end

    test "sets column width" do
      workbook =
        Celex.new()
        |> Celex.add_sheet("Data")
        |> Celex.set_column_width("A", 20)
        |> Celex.put_cell("A1", "Wide column")

      # Column width is stored in the sheet
      sheet = Celex.Workbook.get_sheet(workbook, "Data")
      assert sheet.column_widths["A"] == 20
    end

    test "merges cells" do
      workbook =
        Celex.new()
        |> Celex.add_sheet("Data")
        |> Celex.put_cell("A1", "Merged Header")
        |> Celex.merge_cells("A1:D1")

      sheet = Celex.Workbook.get_sheet(workbook, "Data")
      assert "A1:D1" in sheet.merged_regions
    end

    test "writes workbook to file", %{path: path} do
      workbook =
        Celex.new()
        |> Celex.add_sheet("Sales")
        |> Celex.put_cell("A1", "Product")
        |> Celex.put_cell("B1", "Revenue")
        |> Celex.put_cell("A2", "Widget")
        |> Celex.put_cell("B2", 1500.50)

      assert :ok = Celex.write(workbook, path)
      assert File.exists?(path)
    end

    test "write! raises on error" do
      workbook = Celex.new() |> Celex.add_sheet("Test")
      bad_path = "/nonexistent_#{:rand.uniform(999_999)}/test.xlsx"

      assert_raise RuntimeError, fn ->
        Celex.write!(workbook, bad_path)
      end
    end
  end

  describe "streaming API" do
    test "writes file with stream_write", %{path: path} do
      result =
        Celex.stream_write(path, fn writer ->
          writer
          |> Celex.StreamWriter.start_sheet("Data")
          |> Celex.StreamWriter.write_row(["A", "B", "C"])
          |> Celex.StreamWriter.write_row([1, 2, 3])
        end)

      assert result == :ok
      assert File.exists?(path)
    end

    test "handles stream of rows", %{path: path} do
      rows = Stream.map(1..100, fn i -> ["Row #{i}", i] end)

      result =
        Celex.stream_write(path, fn writer ->
          writer
          |> Celex.StreamWriter.start_sheet("Streamed")
          |> Celex.StreamWriter.write_rows(rows)
        end)

      assert result == :ok
    end

    test "opens file with stream_open", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["A", "B", "C"])
        |> Celex.StreamWriter.write_row([1, 2, 3])
      end)

      {:ok, reader} = Celex.stream_open(path)
      assert Celex.StreamReader.sheet_names(reader) == ["Data"]

      rows = Celex.StreamReader.stream_rows(reader, "Data") |> Enum.to_list()
      assert rows == [["A", "B", "C"], [1, 2, 3]]

      Celex.StreamReader.close(reader)
    end

    test "streams sheet with stream_sheet", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["Name", "Value"])
        |> Celex.StreamWriter.write_row(["A", 1])
        |> Celex.StreamWriter.write_row(["B", 2])
      end)

      rows = Celex.stream_sheet(path, "Data") |> Enum.to_list()
      assert rows == [["Name", "Value"], ["A", 1], ["B", 2]]
    end
  end

  describe "read/1 and read!/1" do
    test "reads an XLSX file", %{path: path} do
      Celex.new()
      |> Celex.add_sheet("Data")
      |> Celex.put_cell("A1", "Hello")
      |> Celex.write!(path)

      {:ok, workbook} = Celex.read(path)
      assert Celex.sheet_names(workbook) == ["Data"]
    end

    test "read! raises on missing file" do
      assert_raise File.Error, fn ->
        Celex.read!("/non/existent/file.xlsx")
      end
    end
  end

  describe "get_sheet/2" do
    test "gets sheet by name", %{path: path} do
      Celex.new()
      |> Celex.add_sheet("Sales")
      |> Celex.put_cell("A1", "Test")
      |> Celex.write!(path)

      workbook = Celex.read!(path)
      {:ok, sheet} = Celex.get_sheet(workbook, "Sales")
      assert sheet.name == "Sales"
    end

    test "returns error for missing sheet", %{path: path} do
      Celex.new()
      |> Celex.add_sheet("Data")
      |> Celex.put_cell("A1", "Test")
      |> Celex.write!(path)

      workbook = Celex.read!(path)
      assert {:error, :sheet_not_found} = Celex.get_sheet(workbook, "Missing")
    end
  end

  describe "get_cell/2 and get_rows/1" do
    test "gets cell values", %{path: path} do
      Celex.new()
      |> Celex.add_sheet("Data")
      |> Celex.put_cell("A1", "Name")
      |> Celex.put_cell("B1", 42)
      |> Celex.write!(path)

      workbook = Celex.read!(path)
      {:ok, sheet} = Celex.get_sheet(workbook, "Data")

      assert Celex.get_cell(sheet, "A1") == "Name"
      assert Celex.get_cell(sheet, "B1") == 42
      assert Celex.get_cell(sheet, {1, 1}) == "Name"
    end

    test "gets all rows", %{path: path} do
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["A", "B"])
        |> Celex.StreamWriter.write_row([1, 2])
      end)

      workbook = Celex.read!(path)
      {:ok, sheet} = Celex.get_sheet(workbook, "Data")

      assert Celex.get_rows(sheet) == [["A", "B"], [1, 2]]
    end
  end

  describe "round-trip" do
    test "data survives write-read cycle", %{path: path} do
      original_data = [
        ["Name", "Age", "Score", "Active"],
        ["Alice", 30, 95.5, true],
        ["Bob", 25, 87.3, false],
        [nil, 42, nil, true]
      ]

      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_rows(original_data)
      end)

      read_data = Celex.stream_sheet(path, "Data") |> Enum.to_list()

      # Compare each row
      assert length(read_data) == 4
      assert Enum.at(read_data, 0) == ["Name", "Age", "Score", "Active"]
      assert Enum.at(read_data, 1) == ["Alice", 30, 95.5, true]
      assert Enum.at(read_data, 2) == ["Bob", 25, 87.3, false]

      # Last row has nil values - check non-nil values match
      last_row = Enum.at(read_data, 3)
      assert Enum.at(last_row, 1) == 42
      assert Enum.at(last_row, 3) == true
    end
  end

  describe "integration" do
    test "full workflow with multiple sheets", %{path: path} do
      workbook =
        Celex.new()
        |> Celex.add_sheet("Summary")
        |> Celex.put_cell("A1", "Summary Report")
        |> Celex.put_cell("A2", "Total Sales")
        |> Celex.put_cell("B2", nil, formula: "SUM(Data!B:B)")
        |> Celex.merge_cells("A1:D1")
        |> Celex.add_sheet("Data")
        |> Celex.put_cell("A1", "Product")
        |> Celex.put_cell("B1", "Amount")
        |> Celex.put_cell("A2", "Widget A")
        |> Celex.put_cell("B2", 100)
        |> Celex.put_cell("A3", "Widget B")
        |> Celex.put_cell("B3", 200)
        |> Celex.set_column_width("A", 20)

      assert Celex.sheet_names(workbook) == ["Summary", "Data"]
      assert :ok = Celex.write(workbook, path)
      assert File.exists?(path)

      # Verify file size is reasonable
      {:ok, %{size: size}} = File.stat(path)
      # Should be at least 1KB
      assert size > 1000
    end
  end
end
