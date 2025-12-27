defmodule Celex.StyleTest do
  use ExUnit.Case, async: true

  alias Celex.Style

  describe "new/0" do
    test "creates empty style with defaults" do
      style = Style.new()
      assert style.font == nil
      assert style.fill == nil
      assert style.border == nil
      assert style.alignment == nil
      assert style.number_format == nil
    end
  end

  describe "new/1 with options" do
    test "creates style with font" do
      style = Style.new(font: %{bold: true, size: 12, name: "Arial"})
      assert style.font.bold == true
      assert style.font.size == 12
      assert style.font.name == "Arial"
    end

    test "creates style with fill" do
      style = Style.new(fill: %{color: "#FF0000", pattern: :solid})
      assert style.fill.color == "#FF0000"
      assert style.fill.pattern == :solid
    end

    test "creates style with border" do
      style = Style.new(border: %{left: :thin, right: :thin, top: :thin, bottom: :thin})
      assert style.border.left == :thin
      assert style.border.right == :thin
    end

    test "creates style with alignment" do
      style = Style.new(alignment: %{horizontal: :center, vertical: :middle, wrap_text: true})
      assert style.alignment.horizontal == :center
      assert style.alignment.vertical == :middle
      assert style.alignment.wrap_text == true
    end

    test "creates style with number format" do
      style = Style.new(number_format: "#,##0.00")
      assert style.number_format == "#,##0.00"
    end
  end

  describe "merge/2" do
    test "merges two styles" do
      style1 = Style.new(font: %{bold: true})
      style2 = Style.new(fill: %{color: "#FF0000"})
      merged = Style.merge(style1, style2)

      assert merged.font.bold == true
      assert merged.fill.color == "#FF0000"
    end

    test "second style overrides first" do
      style1 = Style.new(font: %{bold: true, size: 10})
      style2 = Style.new(font: %{size: 14})
      merged = Style.merge(style1, style2)

      assert merged.font.bold == true
      assert merged.font.size == 14
    end
  end

  describe "Font struct" do
    test "creates font with defaults" do
      font = Style.Font.new()
      assert font.name == nil
      assert font.size == nil
      assert font.bold == false
      assert font.italic == false
      assert font.underline == false
      assert font.color == nil
    end

    test "creates font with options" do
      font = Style.Font.new(bold: true, size: 14, color: "#0000FF")
      assert font.bold == true
      assert font.size == 14
      assert font.color == "#0000FF"
    end
  end

  describe "Fill struct" do
    test "creates fill with defaults" do
      fill = Style.Fill.new()
      assert fill.color == nil
      assert fill.pattern == :none
    end

    test "creates fill with options" do
      fill = Style.Fill.new(color: "#FFFF00", pattern: :solid)
      assert fill.color == "#FFFF00"
      assert fill.pattern == :solid
    end
  end

  describe "Border struct" do
    test "creates border with defaults" do
      border = Style.Border.new()
      assert border.left == nil
      assert border.right == nil
      assert border.top == nil
      assert border.bottom == nil
    end

    test "creates border with options" do
      border = Style.Border.new(left: :thin, right: :medium, color: "#000000")
      assert border.left == :thin
      assert border.right == :medium
      assert border.color == "#000000"
    end

    test "creates border with all sides" do
      border = Style.Border.all(:thin)
      assert border.left == :thin
      assert border.right == :thin
      assert border.top == :thin
      assert border.bottom == :thin
    end
  end

  describe "Alignment struct" do
    test "creates alignment with defaults" do
      alignment = Style.Alignment.new()
      assert alignment.horizontal == nil
      assert alignment.vertical == nil
      assert alignment.wrap_text == false
    end

    test "creates alignment with options" do
      alignment = Style.Alignment.new(horizontal: :center, vertical: :top, wrap_text: true)
      assert alignment.horizontal == :center
      assert alignment.vertical == :top
      assert alignment.wrap_text == true
    end
  end

  describe "predefined styles" do
    test "bold style" do
      style = Style.bold()
      assert style.font.bold == true
    end

    test "italic style" do
      style = Style.italic()
      assert style.font.italic == true
    end

    test "header style" do
      style = Style.header()
      assert style.font.bold == true
      assert style.alignment.horizontal == :center
    end
  end

  describe "number formats" do
    test "currency format" do
      assert Style.number_format(:currency) == ~s("$"#,##0.00)
    end

    test "percentage format" do
      assert Style.number_format(:percentage) == "0.00%"
    end

    test "date format" do
      assert Style.number_format(:date) == "yyyy-mm-dd"
    end

    test "datetime format" do
      assert Style.number_format(:datetime) == "yyyy-mm-dd hh:mm:ss"
    end

    test "custom format passthrough" do
      assert Style.number_format("#,##0.000") == "#,##0.000"
    end
  end
end
