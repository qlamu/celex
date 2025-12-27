defmodule Celex.StreamReaderTest do
  use ExUnit.Case, async: true

  alias Celex.StreamReader

  @test_output_dir System.tmp_dir!()

  # Helper to create a test XLSX file
  defp create_test_xlsx(name, callback) do
    path = Path.join(@test_output_dir, "#{name}_#{:erlang.unique_integer([:positive])}.xlsx")

    :ok = Celex.stream_write(path, callback)

    on_exit(fn -> File.rm(path) end)

    path
  end

  describe "open/1" do
    test "opens an XLSX file" do
      path =
        create_test_xlsx("open", fn writer ->
          writer
          |> Celex.StreamWriter.start_sheet("Sheet1")
          |> Celex.StreamWriter.write_row(["Hello"])
        end)

      assert {:ok, reader} = StreamReader.open(path)
      assert is_struct(reader, StreamReader)
      StreamReader.close(reader)
    end

    test "returns error for non-existent file" do
      assert {:error, _reason} = StreamReader.open("/nonexistent/file.xlsx")
    end
  end

  describe "sheet_names/1" do
    test "returns list of sheet names" do
      path =
        create_test_xlsx("sheets", fn writer ->
          writer
          |> Celex.StreamWriter.start_sheet("Data")
          |> Celex.StreamWriter.write_row(["A"])
          |> Celex.StreamWriter.start_sheet("Summary")
          |> Celex.StreamWriter.write_row(["B"])
        end)

      {:ok, reader} = StreamReader.open(path)
      names = StreamReader.sheet_names(reader)

      assert "Data" in names
      assert "Summary" in names
      StreamReader.close(reader)
    end
  end

  describe "stream_rows/2" do
    test "streams rows from a sheet" do
      path =
        create_test_xlsx("rows", fn writer ->
          writer
          |> Celex.StreamWriter.start_sheet("Data")
          |> Celex.StreamWriter.write_row(["A", "B", "C"])
          |> Celex.StreamWriter.write_row([1, 2, 3])
          |> Celex.StreamWriter.write_row(["X", "Y", "Z"])
        end)

      {:ok, reader} = StreamReader.open(path)
      rows = StreamReader.stream_rows(reader, "Data") |> Enum.to_list()

      assert length(rows) == 3
      assert Enum.at(rows, 0) == ["A", "B", "C"]
      assert Enum.at(rows, 1) == [1, 2, 3]
      assert Enum.at(rows, 2) == ["X", "Y", "Z"]

      StreamReader.close(reader)
    end

    test "streams rows by sheet index" do
      path =
        create_test_xlsx("index", fn writer ->
          writer
          |> Celex.StreamWriter.start_sheet("First")
          |> Celex.StreamWriter.write_row(["First Sheet"])
          |> Celex.StreamWriter.start_sheet("Second")
          |> Celex.StreamWriter.write_row(["Second Sheet"])
        end)

      {:ok, reader} = StreamReader.open(path)

      rows1 = StreamReader.stream_rows(reader, 0) |> Enum.to_list()
      assert Enum.at(rows1, 0) == ["First Sheet"]

      rows2 = StreamReader.stream_rows(reader, 1) |> Enum.to_list()
      assert Enum.at(rows2, 0) == ["Second Sheet"]

      StreamReader.close(reader)
    end

    test "handles various data types" do
      path =
        create_test_xlsx("types", fn writer ->
          writer
          |> Celex.StreamWriter.start_sheet("Types")
          |> Celex.StreamWriter.write_row(["String", 42, 3.14, true, false])
        end)

      {:ok, reader} = StreamReader.open(path)
      [row] = StreamReader.stream_rows(reader, "Types") |> Enum.to_list()

      assert Enum.at(row, 0) == "String"
      assert Enum.at(row, 1) == 42
      assert_in_delta Enum.at(row, 2), 3.14, 0.001
      assert Enum.at(row, 3) == true
      assert Enum.at(row, 4) == false

      StreamReader.close(reader)
    end

    test "handles empty cells" do
      path =
        create_test_xlsx("sparse", fn writer ->
          writer
          |> Celex.StreamWriter.start_sheet("Sparse")
          |> Celex.StreamWriter.write_row(["A", nil, "C"])
        end)

      {:ok, reader} = StreamReader.open(path)
      [row] = StreamReader.stream_rows(reader, "Sparse") |> Enum.to_list()

      assert Enum.at(row, 0) == "A"
      assert Enum.at(row, 1) == nil
      assert Enum.at(row, 2) == "C"

      StreamReader.close(reader)
    end

    test "is lazy - only processes rows as consumed" do
      path =
        create_test_xlsx("lazy", fn writer ->
          rows = for i <- 1..100, do: ["Row #{i}", i]

          writer
          |> Celex.StreamWriter.start_sheet("Data")
          |> Celex.StreamWriter.write_rows(rows)
        end)

      {:ok, reader} = StreamReader.open(path)

      # Take only first 5 rows
      rows = StreamReader.stream_rows(reader, "Data") |> Enum.take(5)

      assert length(rows) == 5
      assert Enum.at(rows, 0) == ["Row 1", 1]

      StreamReader.close(reader)
    end
  end

  describe "read_rows/2" do
    test "reads all rows into a list" do
      path =
        create_test_xlsx("read_all", fn writer ->
          writer
          |> Celex.StreamWriter.start_sheet("Data")
          |> Celex.StreamWriter.write_row(["A", "B"])
          |> Celex.StreamWriter.write_row([1, 2])
        end)

      {:ok, reader} = StreamReader.open(path)
      rows = StreamReader.read_rows(reader, "Data")

      assert rows == [["A", "B"], [1, 2]]

      StreamReader.close(reader)
    end
  end

  describe "integration with large files" do
    test "can read file with many rows" do
      path =
        create_test_xlsx("large", fn writer ->
          rows = Stream.map(1..1000, fn i -> ["Row #{i}", i, i * 1.5] end)

          writer
          |> Celex.StreamWriter.start_sheet("Data")
          |> Celex.StreamWriter.write_rows(rows)
        end)

      {:ok, reader} = StreamReader.open(path)
      count = StreamReader.stream_rows(reader, "Data") |> Enum.count()

      assert count == 1000

      StreamReader.close(reader)
    end
  end
end
