defmodule Celex.Style.Font do
  @moduledoc "Font properties for cell styling."

  @type t :: %__MODULE__{
          name: String.t() | nil,
          size: number() | nil,
          bold: boolean(),
          italic: boolean(),
          underline: boolean(),
          strikethrough: boolean(),
          color: String.t() | nil
        }

  defstruct name: nil,
            size: nil,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            color: nil

  @spec new(keyword()) :: t()
  def new(opts \\ []) do
    struct(__MODULE__, opts)
  end
end

defmodule Celex.Style.Fill do
  @moduledoc "Fill (background) properties for cell styling."

  @type pattern :: :none | :solid | :gray125 | :gray0625

  @type t :: %__MODULE__{
          color: String.t() | nil,
          pattern: pattern()
        }

  defstruct color: nil, pattern: :none

  @spec new(keyword()) :: t()
  def new(opts \\ []) do
    struct(__MODULE__, opts)
  end
end

defmodule Celex.Style.Border do
  @moduledoc "Border properties for cell styling."

  @type border_style :: :none | :thin | :medium | :thick | :dashed | :dotted | :double

  @type t :: %__MODULE__{
          left: border_style() | nil,
          right: border_style() | nil,
          top: border_style() | nil,
          bottom: border_style() | nil,
          diagonal: border_style() | nil,
          color: String.t() | nil
        }

  defstruct left: nil, right: nil, top: nil, bottom: nil, diagonal: nil, color: nil

  @spec new(keyword()) :: t()
  def new(opts \\ []) do
    struct(__MODULE__, opts)
  end

  @doc "Creates a border with all sides set to the same style."
  @spec all(border_style()) :: t()
  def all(style) do
    %__MODULE__{left: style, right: style, top: style, bottom: style}
  end
end

defmodule Celex.Style.Alignment do
  @moduledoc "Alignment properties for cell styling."

  @type horizontal :: :left | :center | :right | :fill | :justify | :distributed
  @type vertical :: :top | :middle | :bottom | :justify | :distributed

  @type t :: %__MODULE__{
          horizontal: horizontal() | nil,
          vertical: vertical() | nil,
          wrap_text: boolean(),
          text_rotation: integer() | nil,
          indent: integer() | nil
        }

  defstruct horizontal: nil, vertical: nil, wrap_text: false, text_rotation: nil, indent: nil

  @spec new(keyword()) :: t()
  def new(opts \\ []) do
    struct(__MODULE__, opts)
  end
end

defmodule Celex.Style do
  @moduledoc """
  Represents cell styling in an Excel worksheet.

  Styles include:
  - Font properties (bold, italic, size, color, name)
  - Fill properties (background color, pattern)
  - Border properties (left, right, top, bottom styles and colors)
  - Alignment properties (horizontal, vertical, wrap text)
  - Number format (currency, percentage, dates, custom formats)

  ## Examples

      # Create a bold style
      style = Style.new(font: %{bold: true, size: 14})

      # Create a style with fill
      style = Style.new(fill: %{color: "#FFFF00", pattern: :solid})

      # Use predefined styles
      Style.bold()
      Style.header()
  """

  alias __MODULE__.{Font, Fill, Border, Alignment}

  @type t :: %__MODULE__{
          font: Font.t() | nil,
          fill: Fill.t() | nil,
          border: Border.t() | nil,
          alignment: Alignment.t() | nil,
          number_format: String.t() | nil
        }

  defstruct [:font, :fill, :border, :alignment, :number_format]

  @doc """
  Creates a new style with the given options.

  ## Options

  - `:font` - Font properties (map or Font struct)
  - `:fill` - Fill properties (map or Fill struct)
  - `:border` - Border properties (map or Border struct)
  - `:alignment` - Alignment properties (map or Alignment struct)
  - `:number_format` - Number format string

  ## Examples

      Style.new()
      Style.new(font: %{bold: true})
      Style.new(font: %{bold: true}, fill: %{color: "#FF0000"})
  """
  @spec new(keyword()) :: t()
  def new(opts \\ []) do
    %__MODULE__{
      font: build_font(Keyword.get(opts, :font)),
      fill: build_fill(Keyword.get(opts, :fill)),
      border: build_border(Keyword.get(opts, :border)),
      alignment: build_alignment(Keyword.get(opts, :alignment)),
      number_format: Keyword.get(opts, :number_format)
    }
  end

  @doc """
  Merges two styles, with the second style taking precedence.
  """
  @spec merge(t(), t()) :: t()
  def merge(%__MODULE__{} = style1, %__MODULE__{} = style2) do
    %__MODULE__{
      font: merge_component(style1.font, style2.font, Font),
      fill: merge_component(style1.fill, style2.fill, Fill),
      border: merge_component(style1.border, style2.border, Border),
      alignment: merge_component(style1.alignment, style2.alignment, Alignment),
      number_format: style2.number_format || style1.number_format
    }
  end

  defp merge_component(nil, nil, _module), do: nil
  defp merge_component(nil, comp2, _module), do: comp2
  defp merge_component(comp1, nil, _module), do: comp1

  defp merge_component(comp1, comp2, module) do
    map1 = Map.from_struct(comp1)
    map2 = Map.from_struct(comp2)
    defaults = Map.from_struct(struct(module))

    merged =
      Map.merge(map1, map2, fn key, v1, v2 ->
        # Only use v2 if it's different from the default
        # This allows the first style's value to persist when second doesn't set it
        if v2 == Map.get(defaults, key), do: v1, else: v2
      end)

    struct(module, merged)
  end

  # Predefined styles

  @doc "Creates a bold style."
  @spec bold() :: t()
  def bold, do: new(font: %{bold: true})

  @doc "Creates an italic style."
  @spec italic() :: t()
  def italic, do: new(font: %{italic: true})

  @doc "Creates a header style (bold, centered)."
  @spec header() :: t()
  def header do
    new(
      font: %{bold: true},
      alignment: %{horizontal: :center}
    )
  end

  # Number format helpers

  @doc """
  Returns a number format string for common formats.

  ## Supported atoms

  - `:currency` - "$"#,##0.00
  - `:percentage` - 0.00%
  - `:date` - yyyy-mm-dd
  - `:datetime` - yyyy-mm-dd hh:mm:ss
  - `:time` - hh:mm:ss
  - `:integer` - #,##0
  - `:decimal` - #,##0.00

  Custom format strings are passed through unchanged.
  """
  @spec number_format(atom() | String.t()) :: String.t()
  def number_format(:currency), do: ~s("$"#,##0.00)
  def number_format(:percentage), do: "0.00%"
  def number_format(:date), do: "yyyy-mm-dd"
  def number_format(:datetime), do: "yyyy-mm-dd hh:mm:ss"
  def number_format(:time), do: "hh:mm:ss"
  def number_format(:integer), do: "#,##0"
  def number_format(:decimal), do: "#,##0.00"
  def number_format(custom) when is_binary(custom), do: custom

  # Private builders

  defp build_font(nil), do: nil
  defp build_font(%Font{} = font), do: font
  defp build_font(map) when is_map(map), do: Font.new(Map.to_list(map))

  defp build_fill(nil), do: nil
  defp build_fill(%Fill{} = fill), do: fill
  defp build_fill(map) when is_map(map), do: Fill.new(Map.to_list(map))

  defp build_border(nil), do: nil
  defp build_border(%Border{} = border), do: border
  defp build_border(map) when is_map(map), do: Border.new(Map.to_list(map))

  defp build_alignment(nil), do: nil
  defp build_alignment(%Alignment{} = alignment), do: alignment
  defp build_alignment(map) when is_map(map), do: Alignment.new(Map.to_list(map))
end
