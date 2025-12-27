defmodule Celex.XML.SharedStringsTest do
  use ExUnit.Case, async: true

  alias Celex.XML.SharedStrings

  describe "generate/1" do
    test "generates valid shared strings XML" do
      xml = SharedStrings.generate(["Hello", "World"])

      assert String.starts_with?(xml, "<?xml")
      assert String.contains?(xml, "<sst")

      assert String.contains?(
               xml,
               "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
             )
    end

    test "includes string items" do
      xml = SharedStrings.generate(["Hello", "World"])

      assert String.contains?(xml, "<si><t>Hello</t></si>")
      assert String.contains?(xml, "<si><t>World</t></si>")
    end

    test "includes count and unique count" do
      xml = SharedStrings.generate(["A", "B", "C"])

      assert String.contains?(xml, "count=\"3\"")
      assert String.contains?(xml, "uniqueCount=\"3\"")
    end

    test "escapes special characters" do
      xml = SharedStrings.generate(["Data & More", "<test>"])

      assert String.contains?(xml, "Data &amp; More")
      assert String.contains?(xml, "&lt;test&gt;")
    end

    test "handles empty list" do
      xml = SharedStrings.generate([])

      assert String.contains?(xml, "count=\"0\"")
      assert String.contains?(xml, "uniqueCount=\"0\"")
    end

    test "preserves whitespace in strings" do
      xml = SharedStrings.generate(["  spaced  "])

      assert String.contains?(xml, "<t xml:space=\"preserve\">  spaced  </t>")
    end
  end

  describe "build_table/1" do
    test "builds string to index mapping" do
      {table, _} = SharedStrings.build_table(["A", "B", "C", "A", "B"])

      assert table["A"] == 0
      assert table["B"] == 1
      assert table["C"] == 2
    end

    test "returns unique strings list" do
      {_, strings} = SharedStrings.build_table(["A", "B", "C", "A", "B"])

      assert strings == ["A", "B", "C"]
    end
  end
end
