defmodule CelexStyleTest do
  use ExUnit.Case
  alias Celex.Cell

  setup do
    File.mkdir_p("test_output")
    :ok
  end

  test "applies all styling options correctly" do
    filename = "test_output/all_styles.xlsx"
    data = [
      [
        Cell.new("Bold", bold: true),
        Cell.new("Italic", italic: true),
        Cell.new("Underline", underline: true),
        Cell.new("Red Font", font_color: "FF0000"),
        Cell.new("Blue Background", background_color: "0000FF"),
        Cell.new("Right Aligned", alignment: "right"),
        Cell.new(123, number_format: "0.0")
      ]
    ]

    assert {:ok, ^filename} = Celex.create_excel(filename, data)
    assert File.exists?(filename)

    {:ok, files} = :zip.unzip(String.to_charlist(filename), [:memory])
    {_, styles_xml_binary} = Enum.find(files, fn {path, _} -> path == ~c"xl/styles.xml" end)
    styles_xml = to_string(styles_xml_binary)

    # Verify font styles
    assert String.contains?(styles_xml, "<b/>")
    assert String.contains?(styles_xml, "<i/>")
    assert String.contains?(styles_xml, "<u/>")
    assert String.contains?(styles_xml, ~s(<color rgb="FF0000"/>))

    # Verify fill style
    assert String.contains?(styles_xml, ~s(<fgColor rgb="0000FF"/>))

    # Verify alignment
    assert String.contains?(styles_xml, ~s(<alignment horizontal="right"/>))

    # Verify number format
    assert String.contains?(styles_xml, ~s(<numFmt numFmtId="164" formatCode="0.0"/>))
  end

  test "handles mixed styled and unstyled cells" do
    filename = "test_output/mixed_styles.xlsx"
    data = [
      ["Plain", Cell.new("Styled", bold: true), "Plain Again"]
    ]

    assert {:ok, ^filename} = Celex.create_excel(filename, data)
    assert File.exists?(filename)

    {:ok, files} = :zip.unzip(String.to_charlist(filename), [:memory])
    {_, styles_xml_binary} = Enum.find(files, fn {path, _} -> path == ~c"xl/styles.xml" end)
    styles_xml = to_string(styles_xml_binary)

    assert String.contains?(styles_xml, "<b/>")
  end

  test "handles empty input" do
    filename = "test_output/empty.xlsx"
    assert {:ok, ^filename} = Celex.create_excel(filename, [])
    assert File.exists?(filename)
  end

  test "handles empty rows and nil values" do
    filename = "test_output/empty_rows.xlsx"
    data = [
      [1, 2, 3],
      [],
      [4, nil, 6]
    ]

    assert {:ok, ^filename} = Celex.create_excel(filename, data)
    assert File.exists?(filename)
  end
end