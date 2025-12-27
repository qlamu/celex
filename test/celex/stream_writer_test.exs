defmodule Celex.StreamWriterTest do
  use ExUnit.Case, async: true

  alias Celex.StreamWriter

  @test_output_dir System.tmp_dir!()

  setup do
    # Create a unique filename for each test
    filename = "celex_test_#{:erlang.unique_integer([:positive])}.xlsx"
    path = Path.join(@test_output_dir, filename)

    on_exit(fn ->
      File.rm(path)
    end)

    {:ok, path: path}
  end

  describe "open/1" do
    test "creates a new stream writer", %{path: path} do
      {:ok, writer} = StreamWriter.open(path)
      assert writer.path == path
      assert writer.sheets == []
      StreamWriter.abort(writer)
    end
  end

  describe "start_sheet/2" do
    test "starts a new sheet", %{path: path} do
      {:ok, writer} = StreamWriter.open(path)
      writer = StreamWriter.start_sheet(writer, "Data")

      assert writer.current_sheet == "Data"
      assert "Data" in writer.sheets

      StreamWriter.abort(writer)
    end

    test "can start multiple sheets", %{path: path} do
      {:ok, writer} = StreamWriter.open(path)

      writer =
        writer
        |> StreamWriter.start_sheet("Sheet1")
        |> StreamWriter.write_row(["A1"])
        |> StreamWriter.start_sheet("Sheet2")
        |> StreamWriter.write_row(["B1"])

      assert length(writer.sheets) == 2
      StreamWriter.abort(writer)
    end
  end

  describe "write_row/2" do
    test "writes a row of values", %{path: path} do
      {:ok, writer} = StreamWriter.open(path)

      writer =
        writer
        |> StreamWriter.start_sheet("Data")
        |> StreamWriter.write_row(["A", "B", "C"])
        |> StreamWriter.write_row([1, 2, 3])

      assert writer.row_count == 2
      StreamWriter.abort(writer)
    end

    test "handles various data types", %{path: path} do
      {:ok, writer} = StreamWriter.open(path)

      writer =
        writer
        |> StreamWriter.start_sheet("Data")
        |> StreamWriter.write_row(["String", 42, 3.14, true, false, ~D[2024-06-15], nil])

      assert writer.row_count == 1
      StreamWriter.abort(writer)
    end
  end

  describe "write_rows/2" do
    test "writes multiple rows from enumerable", %{path: path} do
      {:ok, writer} = StreamWriter.open(path)

      rows = [
        ["A1", "B1", "C1"],
        ["A2", "B2", "C2"],
        ["A3", "B3", "C3"]
      ]

      writer =
        writer
        |> StreamWriter.start_sheet("Data")
        |> StreamWriter.write_rows(rows)

      assert writer.row_count == 3
      StreamWriter.abort(writer)
    end

    test "writes from stream lazily", %{path: path} do
      {:ok, writer} = StreamWriter.open(path)

      stream = Stream.map(1..100, fn i -> ["Row #{i}", i] end)

      writer =
        writer
        |> StreamWriter.start_sheet("Data")
        |> StreamWriter.write_rows(stream)

      assert writer.row_count == 100
      StreamWriter.abort(writer)
    end
  end

  describe "close/1" do
    test "creates valid XLSX file", %{path: path} do
      {:ok, writer} = StreamWriter.open(path)

      :ok =
        writer
        |> StreamWriter.start_sheet("Data")
        |> StreamWriter.write_row(["Hello", "World"])
        |> StreamWriter.write_row([1, 2])
        |> StreamWriter.close()

      assert File.exists?(path)

      # Verify it's a valid ZIP
      {:ok, files} = :zip.list_dir(String.to_charlist(path))

      file_names =
        files
        |> Enum.filter(fn
          {:zip_file, _, _, _, _, _} -> true
          _ -> false
        end)
        |> Enum.map(fn {:zip_file, name, _, _, _, _} -> to_string(name) end)

      assert "[Content_Types].xml" in file_names
      assert "_rels/.rels" in file_names
      assert "xl/workbook.xml" in file_names
      assert "xl/worksheets/sheet1.xml" in file_names
    end

    test "creates file with multiple sheets", %{path: path} do
      {:ok, writer} = StreamWriter.open(path)

      :ok =
        writer
        |> StreamWriter.start_sheet("Sheet1")
        |> StreamWriter.write_row(["Sheet 1 Data"])
        |> StreamWriter.start_sheet("Sheet2")
        |> StreamWriter.write_row(["Sheet 2 Data"])
        |> StreamWriter.close()

      {:ok, files} = :zip.list_dir(String.to_charlist(path))

      file_names =
        files
        |> Enum.filter(fn
          {:zip_file, _, _, _, _, _} -> true
          _ -> false
        end)
        |> Enum.map(fn {:zip_file, name, _, _, _, _} -> to_string(name) end)

      assert "xl/worksheets/sheet1.xml" in file_names
      assert "xl/worksheets/sheet2.xml" in file_names
    end
  end

  describe "abort/1" do
    test "cleans up temp files", %{path: path} do
      {:ok, writer} = StreamWriter.open(path)

      writer =
        writer
        |> StreamWriter.start_sheet("Data")
        |> StreamWriter.write_row(["Test"])

      :ok = StreamWriter.abort(writer)

      # Should not create the output file
      refute File.exists?(path)
    end
  end

  describe "integration" do
    test "can write large number of rows", %{path: path} do
      {:ok, writer} = StreamWriter.open(path)

      # Write 10,000 rows
      rows = Stream.map(1..10_000, fn i -> ["Row #{i}", i, i * 1.5] end)

      :ok =
        writer
        |> StreamWriter.start_sheet("LargeData")
        |> StreamWriter.write_rows(rows)
        |> StreamWriter.close()

      assert File.exists?(path)
      # File should be reasonably sized
      {:ok, %{size: size}} = File.stat(path)
      # At least 100KB
      assert size > 100_000
    end
  end
end
