defmodule CelexTest do
  use ExUnit.Case
  doctest Celex

  alias Celex

  describe "excel_column_name/1" do
    test "converts single digit column indices to letters" do
      assert Celex.excel_column_name(0) == "A"
      assert Celex.excel_column_name(1) == "B"
      assert Celex.excel_column_name(25) == "Z"
    end

    test "converts two digit column indices to letters" do
      assert Celex.excel_column_name(26) == "AA"
      assert Celex.excel_column_name(27) == "AB"
      assert Celex.excel_column_name(51) == "AZ"
      assert Celex.excel_column_name(52) == "BA"
      assert Celex.excel_column_name(701) == "ZZ"
    end

    test "converts three digit column indices to letters" do
      assert Celex.excel_column_name(702) == "AAA"
      assert Celex.excel_column_name(703) == "AAB"
      assert Celex.excel_column_name(728) == "ABA"
    end

    test "handles edge cases" do
      assert Celex.excel_column_name(675) == "YZ"
      assert Celex.excel_column_name(676) == "ZA"
      assert Celex.excel_column_name(16383) == "XFD"
    end
  end
end

