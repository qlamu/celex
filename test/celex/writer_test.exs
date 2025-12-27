defmodule Celex.WriterTest do
  use ExUnit.Case, async: true

  alias Celex.Writer
  alias Celex.Workbook
  alias Celex.Sheet

  @test_output_dir System.tmp_dir!()

  setup do
    filename = "celex_writer_test_#{:erlang.unique_integer([:positive])}.xlsx"
    path = Path.join(@test_output_dir, filename)

    on_exit(fn ->
      File.rm(path)
    end)

    {:ok, path: path}
  end

  describe "write/2" do
    test "writes a workbook to file", %{path: path} do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Data")
        |> Workbook.update_sheet("Data", fn sheet ->
          sheet
          |> Sheet.put_cell("A1", "Hello")
          |> Sheet.put_cell("B1", "World")
        end)

      assert :ok = Writer.write(workbook, path)
      assert File.exists?(path)
    end

    test "writes multiple sheets", %{path: path} do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Sheet1")
        |> Workbook.update_sheet("Sheet1", fn sheet ->
          Sheet.put_cell(sheet, "A1", "Sheet 1")
        end)
        |> Workbook.add_sheet("Sheet2")
        |> Workbook.update_sheet("Sheet2", fn sheet ->
          Sheet.put_cell(sheet, "A1", "Sheet 2")
        end)

      assert :ok = Writer.write(workbook, path)

      # Verify both sheets exist
      {:ok, files} = :zip.list_dir(String.to_charlist(path))

      file_names =
        files
        |> Enum.filter(&match?({:zip_file, _, _, _, _, _}, &1))
        |> Enum.map(fn {:zip_file, name, _, _, _, _} -> to_string(name) end)

      assert "xl/worksheets/sheet1.xml" in file_names
      assert "xl/worksheets/sheet2.xml" in file_names
    end

    test "preserves various data types", %{path: path} do
      workbook =
        Workbook.new()
        |> Workbook.add_sheet("Types")
        |> Workbook.update_sheet("Types", fn sheet ->
          sheet
          |> Sheet.put_cell("A1", "String")
          |> Sheet.put_cell("A2", 42)
          |> Sheet.put_cell("A3", 3.14)
          |> Sheet.put_cell("A4", true)
          |> Sheet.put_cell("A5", false)
          |> Sheet.put_cell("A6", ~D[2024-06-15])
        end)

      assert :ok = Writer.write(workbook, path)
      assert File.exists?(path)
    end

    test "handles empty workbook", %{path: path} do
      workbook = Workbook.new() |> Workbook.add_sheet("Empty")

      assert :ok = Writer.write(workbook, path)
      assert File.exists?(path)
    end
  end

  describe "write!/2" do
    test "raises on error" do
      workbook = Workbook.new() |> Workbook.add_sheet("Test")
      # Try to write to a non-existent directory
      bad_path = "/nonexistent_dir_#{:rand.uniform(999_999)}/test.xlsx"

      assert_raise RuntimeError, fn ->
        Writer.write!(workbook, bad_path)
      end
    end
  end
end
