defmodule CelexTest do
  use ExUnit.Case
  doctest Celex

  alias Celex.{Cell, Worksheet}

  setup do
    # Create a directory for test output
    File.mkdir_p("test_output")
    :ok
  end

  # Test basic file creation with a list of lists
  test "creates a valid XLSX file from a list of lists" do
    filename = "test_output/simple_list.xlsx"
    data = [[1, 2, 3], [4, 5, 6]]

    assert {:ok, ^filename} = Celex.create_excel(filename, data)
    assert File.exists?(filename)
  end

  # Test file creation with a map of sheet names to data
  test "creates a valid XLSX file from a map of data" do
    filename = "test_output/map_data.xlsx"
    data = %{
      "Sheet1" => [[1, 2], [3, 4]],
      "Sheet2" => [["a", "b"], ["c", "d"]]
    }

    assert {:ok, ^filename} = Celex.create_excel(filename, data)
    assert File.exists?(filename)
  end

  # Test file creation with a list of Worksheet structs
  test "creates a valid XLSX file from a list of Worksheet structs" do
    filename = "test_output/worksheet_structs.xlsx"
    worksheets = [
      Worksheet.new("DataSheet", [[1, 2, 3]]),
      Worksheet.new("TextSheet", [["a", "b", "c"]])
    ]

    assert {:ok, ^filename} = Celex.create_excel(filename, worksheets)
    assert File.exists?(filename)
  end

  # Test handling of different cell types
  test "correctly handles numbers, strings, and formulas" do
    filename = "test_output/cell_types.xlsx"
    data = [
      [10, "hello", "=SUM(A1:A1)"],
      [3.14, "world", "=A2"]
    ]

    assert {:ok, ^filename} = Celex.create_excel(filename, data)
    assert File.exists?(filename)
  end

  # Test that a custom number format can be applied
  test "applies a custom number format and generates correct styles" do
    filename = "test_output/custom_number_format.xlsx"
    data = [
      [
        Cell.new("Bold Text", bold: true),
        Cell.new(0.123, number_format: "0.00%")
      ]
    ]

    assert {:ok, ^filename} = Celex.create_excel(filename, data)
    assert File.exists?(filename)

    # Unzip and inspect the styles.xml file to verify correctness
    {:ok, files} = :zip.unzip(String.to_charlist(filename), [:memory])
    {_, styles_xml_binary} = Enum.find(files, fn {path, _} -> path == ~c"xl/styles.xml" end)
    styles_xml = to_string(styles_xml_binary)

    # Assert that the custom number format is defined with the correct ID (164)
    assert String.contains?(styles_xml, ~s(<numFmt numFmtId="164" formatCode="0.00%"/>))

    # Assert that the cell's xf record uses the correct numFmtId
    assert String.contains?(styles_xml, ~s(<xf numFmtId="164"))
  end
end