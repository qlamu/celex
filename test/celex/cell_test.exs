defmodule Celex.CellTest do
  use ExUnit.Case, async: true

  alias Celex.Cell

  describe "new/1" do
    test "creates cell with string value" do
      cell = Cell.new("Hello")
      assert cell.value == "Hello"
      assert cell.type == :string
      assert cell.style == nil
      assert cell.formula == nil
    end

    test "creates cell with integer value" do
      cell = Cell.new(42)
      assert cell.value == 42
      assert cell.type == :number
    end

    test "creates cell with float value" do
      cell = Cell.new(3.14159)
      assert cell.value == 3.14159
      assert cell.type == :number
    end

    test "creates cell with boolean true" do
      cell = Cell.new(true)
      assert cell.value == true
      assert cell.type == :boolean
    end

    test "creates cell with boolean false" do
      cell = Cell.new(false)
      assert cell.value == false
      assert cell.type == :boolean
    end

    test "creates cell with date" do
      date = ~D[2024-06-15]
      cell = Cell.new(date)
      assert cell.value == date
      assert cell.type == :date
    end

    test "creates cell with datetime" do
      datetime = ~N[2024-06-15 14:30:00]
      cell = Cell.new(datetime)
      assert cell.value == datetime
      assert cell.type == :datetime
    end

    test "creates cell with time" do
      time = ~T[14:30:00]
      cell = Cell.new(time)
      assert cell.value == time
      assert cell.type == :time
    end

    test "creates cell with nil value" do
      cell = Cell.new(nil)
      assert cell.value == nil
      assert cell.type == :blank
    end

    test "creates cell with empty string" do
      cell = Cell.new("")
      assert cell.value == ""
      assert cell.type == :string
    end
  end

  describe "new/2 with options" do
    test "creates cell with style" do
      style = %{bold: true, font_size: 12}
      cell = Cell.new("Hello", style: style)
      assert cell.value == "Hello"
      assert cell.style == style
    end

    test "creates cell with formula" do
      cell = Cell.new(nil, formula: "SUM(A1:A10)")
      assert cell.formula == "SUM(A1:A10)"
      assert cell.type == :formula
    end

    test "creates cell with formula and cached value" do
      cell = Cell.new(100, formula: "SUM(A1:A10)")
      assert cell.value == 100
      assert cell.formula == "SUM(A1:A10)"
      assert cell.type == :formula
    end
  end

  describe "formula/1" do
    test "creates formula cell from string starting with =" do
      cell = Cell.formula("=SUM(A1:A10)")
      assert cell.formula == "SUM(A1:A10)"
      assert cell.type == :formula
      assert cell.value == nil
    end

    test "creates formula cell from string without =" do
      cell = Cell.formula("SUM(A1:A10)")
      assert cell.formula == "SUM(A1:A10)"
      assert cell.type == :formula
    end
  end

  describe "blank?/1" do
    test "returns true for nil value" do
      cell = Cell.new(nil)
      assert Cell.blank?(cell) == true
    end

    test "returns false for non-nil value" do
      cell = Cell.new("Hello")
      assert Cell.blank?(cell) == false
    end

    test "returns false for empty string" do
      cell = Cell.new("")
      assert Cell.blank?(cell) == false
    end

    test "returns false for zero" do
      cell = Cell.new(0)
      assert Cell.blank?(cell) == false
    end

    test "returns false for false boolean" do
      cell = Cell.new(false)
      assert Cell.blank?(cell) == false
    end
  end

  describe "formula?/1" do
    test "returns true for formula cell" do
      cell = Cell.new(nil, formula: "SUM(A1:A10)")
      assert Cell.formula?(cell) == true
    end

    test "returns false for non-formula cell" do
      cell = Cell.new("Hello")
      assert Cell.formula?(cell) == false
    end
  end

  describe "type detection" do
    test "detects type from value automatically" do
      assert Cell.new("text").type == :string
      assert Cell.new(123).type == :number
      assert Cell.new(1.5).type == :number
      assert Cell.new(true).type == :boolean
      assert Cell.new(false).type == :boolean
      assert Cell.new(~D[2024-01-01]).type == :date
      assert Cell.new(~N[2024-01-01 12:00:00]).type == :datetime
      assert Cell.new(~T[12:00:00]).type == :time
      assert Cell.new(nil).type == :blank
    end
  end
end
