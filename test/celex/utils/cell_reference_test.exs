defmodule Celex.Utils.CellReferenceTest do
  use ExUnit.Case, async: true

  alias Celex.Utils.CellReference

  describe "parse/1" do
    test "parses simple cell references" do
      assert CellReference.parse("A1") == {:ok, {1, 1}}
      assert CellReference.parse("B1") == {:ok, {1, 2}}
      assert CellReference.parse("A2") == {:ok, {2, 1}}
      assert CellReference.parse("C3") == {:ok, {3, 3}}
    end

    test "parses double letter columns" do
      assert CellReference.parse("AA1") == {:ok, {1, 27}}
      assert CellReference.parse("AB1") == {:ok, {1, 28}}
      assert CellReference.parse("AZ1") == {:ok, {1, 52}}
      assert CellReference.parse("BA1") == {:ok, {1, 53}}
    end

    test "parses triple letter columns" do
      assert CellReference.parse("AAA1") == {:ok, {1, 703}}
      # Excel max column
      assert CellReference.parse("XFD1") == {:ok, {1, 16384}}
    end

    test "parses large row numbers" do
      # Excel max row
      assert CellReference.parse("A1048576") == {:ok, {1_048_576, 1}}
      assert CellReference.parse("XFD1048576") == {:ok, {1_048_576, 16384}}
    end

    test "handles lowercase input" do
      assert CellReference.parse("a1") == {:ok, {1, 1}}
      assert CellReference.parse("aa1") == {:ok, {1, 27}}
      assert CellReference.parse("Ab1") == {:ok, {1, 28}}
    end

    test "returns error for invalid references" do
      assert CellReference.parse("") == {:error, :invalid_reference}
      assert CellReference.parse("1A") == {:error, :invalid_reference}
      assert CellReference.parse("A") == {:error, :invalid_reference}
      assert CellReference.parse("1") == {:error, :invalid_reference}
      assert CellReference.parse("A0") == {:error, :invalid_reference}
      assert CellReference.parse("A-1") == {:error, :invalid_reference}
    end
  end

  describe "to_string/1" do
    test "converts row/col tuple to cell reference" do
      assert CellReference.to_string({1, 1}) == "A1"
      assert CellReference.to_string({1, 2}) == "B1"
      assert CellReference.to_string({2, 1}) == "A2"
      assert CellReference.to_string({3, 3}) == "C3"
    end

    test "converts double letter columns" do
      assert CellReference.to_string({1, 27}) == "AA1"
      assert CellReference.to_string({1, 28}) == "AB1"
      assert CellReference.to_string({1, 52}) == "AZ1"
      assert CellReference.to_string({1, 53}) == "BA1"
    end

    test "converts triple letter columns" do
      assert CellReference.to_string({1, 703}) == "AAA1"
      assert CellReference.to_string({1, 16384}) == "XFD1"
    end

    test "handles large row numbers" do
      assert CellReference.to_string({1_048_576, 1}) == "A1048576"
    end
  end

  describe "column_to_index/1" do
    test "converts column letter to 1-based index" do
      assert CellReference.column_to_index("A") == 1
      assert CellReference.column_to_index("B") == 2
      assert CellReference.column_to_index("Z") == 26
      assert CellReference.column_to_index("AA") == 27
      assert CellReference.column_to_index("AB") == 28
      assert CellReference.column_to_index("AZ") == 52
      assert CellReference.column_to_index("BA") == 53
      assert CellReference.column_to_index("ZZ") == 702
      assert CellReference.column_to_index("AAA") == 703
    end

    test "handles lowercase" do
      assert CellReference.column_to_index("a") == 1
      assert CellReference.column_to_index("aa") == 27
    end
  end

  describe "index_to_column/1" do
    test "converts 1-based index to column letter" do
      assert CellReference.index_to_column(1) == "A"
      assert CellReference.index_to_column(2) == "B"
      assert CellReference.index_to_column(26) == "Z"
      assert CellReference.index_to_column(27) == "AA"
      assert CellReference.index_to_column(28) == "AB"
      assert CellReference.index_to_column(52) == "AZ"
      assert CellReference.index_to_column(53) == "BA"
      assert CellReference.index_to_column(702) == "ZZ"
      assert CellReference.index_to_column(703) == "AAA"
      assert CellReference.index_to_column(16384) == "XFD"
    end
  end

  describe "parse_range/1" do
    test "parses cell range" do
      assert CellReference.parse_range("A1:B2") == {:ok, {{1, 1}, {2, 2}}}
      assert CellReference.parse_range("A1:Z100") == {:ok, {{1, 1}, {100, 26}}}
      assert CellReference.parse_range("AA1:AB10") == {:ok, {{1, 27}, {10, 28}}}
    end

    test "handles single cell as range" do
      assert CellReference.parse_range("A1:A1") == {:ok, {{1, 1}, {1, 1}}}
    end

    test "returns error for invalid ranges" do
      assert CellReference.parse_range("A1") == {:error, :invalid_range}
      assert CellReference.parse_range("A1:") == {:error, :invalid_range}
      assert CellReference.parse_range(":B2") == {:error, :invalid_range}
      assert CellReference.parse_range("") == {:error, :invalid_range}
    end
  end

  describe "roundtrip" do
    test "parse and to_string are inverse operations" do
      for col <- [1, 26, 27, 52, 100, 702, 703, 16384],
          row <- [1, 100, 1000, 1_048_576] do
        ref = CellReference.to_string({row, col})
        assert CellReference.parse(ref) == {:ok, {row, col}}
      end
    end

    test "column_to_index and index_to_column are inverse operations" do
      for col <- 1..1000 do
        letters = CellReference.index_to_column(col)
        assert CellReference.column_to_index(letters) == col
      end
    end
  end
end
