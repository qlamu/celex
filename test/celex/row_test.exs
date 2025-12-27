defmodule Celex.RowTest do
  use ExUnit.Case, async: true

  alias Celex.Row
  alias Celex.Cell

  describe "new/1" do
    test "creates row from list of values" do
      row = Row.new(["A", "B", "C"])
      assert row.index == nil
      assert length(row.cells) == 3
      assert Enum.at(row.cells, 0).value == "A"
      assert Enum.at(row.cells, 1).value == "B"
      assert Enum.at(row.cells, 2).value == "C"
    end

    test "creates row from list of mixed types" do
      row = Row.new(["Hello", 42, true, ~D[2024-06-15], nil])
      assert length(row.cells) == 5
      assert Enum.at(row.cells, 0).type == :string
      assert Enum.at(row.cells, 1).type == :number
      assert Enum.at(row.cells, 2).type == :boolean
      assert Enum.at(row.cells, 3).type == :date
      assert Enum.at(row.cells, 4).type == :blank
    end

    test "creates row from list of Cell structs" do
      cells = [Cell.new("A"), Cell.new(1), Cell.new(true)]
      row = Row.new(cells)
      assert length(row.cells) == 3
      assert Enum.at(row.cells, 0).value == "A"
    end

    test "creates row from mixed values and Cell structs" do
      row = Row.new(["A", Cell.new(1, style: %{bold: true}), "C"])
      assert length(row.cells) == 3
      assert Enum.at(row.cells, 1).style != nil
    end

    test "creates empty row" do
      row = Row.new([])
      assert row.cells == []
    end
  end

  describe "new/2 with options" do
    test "creates row with index" do
      row = Row.new(["A", "B"], index: 5)
      assert row.index == 5
    end

    test "creates row with height" do
      row = Row.new(["A", "B"], height: 20.5)
      assert row.height == 20.5
    end

    test "creates row with style applied to all cells" do
      row = Row.new(["A", "B", "C"], style: %{bold: true})
      assert Enum.all?(row.cells, fn cell -> cell.style == %{bold: true} end)
    end
  end

  describe "from_map/2" do
    test "creates row from map with column order" do
      data = %{name: "Alice", age: 30, city: "NYC"}
      columns = [:name, :age, :city]
      row = Row.from_map(data, columns)

      assert length(row.cells) == 3
      assert Enum.at(row.cells, 0).value == "Alice"
      assert Enum.at(row.cells, 1).value == 30
      assert Enum.at(row.cells, 2).value == "NYC"
    end

    test "handles missing keys with nil" do
      data = %{name: "Alice"}
      columns = [:name, :age, :city]
      row = Row.from_map(data, columns)

      assert Enum.at(row.cells, 0).value == "Alice"
      assert Enum.at(row.cells, 1).value == nil
      assert Enum.at(row.cells, 2).value == nil
    end
  end

  describe "values/1" do
    test "extracts values from row" do
      row = Row.new(["A", 1, true])
      assert Row.values(row) == ["A", 1, true]
    end

    test "returns empty list for empty row" do
      row = Row.new([])
      assert Row.values(row) == []
    end
  end

  describe "size/1" do
    test "returns number of cells" do
      row = Row.new(["A", "B", "C"])
      assert Row.size(row) == 3
    end

    test "returns 0 for empty row" do
      row = Row.new([])
      assert Row.size(row) == 0
    end
  end

  describe "get/2" do
    test "gets cell by 0-based index" do
      row = Row.new(["A", "B", "C"])
      assert Row.get(row, 0).value == "A"
      assert Row.get(row, 1).value == "B"
      assert Row.get(row, 2).value == "C"
    end

    test "returns nil for out of bounds" do
      row = Row.new(["A", "B"])
      assert Row.get(row, 5) == nil
    end
  end

  describe "put/3" do
    test "puts cell at index" do
      row = Row.new(["A", "B", "C"])
      updated = Row.put(row, 1, "X")
      assert Row.get(updated, 1).value == "X"
    end

    test "expands row if index is beyond current size" do
      row = Row.new(["A"])
      updated = Row.put(row, 3, "D")
      assert Row.size(updated) == 4
      assert Row.get(updated, 3).value == "D"
      assert Row.get(updated, 1).value == nil
      assert Row.get(updated, 2).value == nil
    end
  end
end
