defmodule Celex.Cell do
  @moduledoc """
  Represents a single cell in an Excel worksheet.

  A cell contains:
  - `value` - The cell's value (string, number, boolean, date, datetime, time, or nil)
  - `type` - The inferred type (:string, :number, :boolean, :date, :datetime, :time, :blank, :formula)
  - `style` - Optional style information (font, fill, border, alignment, number format)
  - `formula` - Optional formula string (without the leading '=')

  ## Examples

      # Simple values
      Cell.new("Hello")
      Cell.new(42)
      Cell.new(~D[2024-06-15])

      # With style
      Cell.new("Header", style: %{bold: true})

      # With formula
      Cell.new(100, formula: "SUM(A1:A10)")

      # Formula only (value will be calculated by Excel)
      Cell.formula("=SUM(A1:A10)")
  """

  @type cell_type :: :string | :number | :boolean | :date | :datetime | :time | :blank | :formula
  @type value :: String.t() | number() | boolean() | Date.t() | NaiveDateTime.t() | Time.t() | nil

  @type t :: %__MODULE__{
          value: value(),
          type: cell_type(),
          style: map() | nil,
          formula: String.t() | nil
        }

  defstruct [:value, :type, :style, :formula]

  @doc """
  Creates a new cell with the given value.

  The cell type is automatically inferred from the value:
  - String -> :string
  - Integer/Float -> :number
  - Boolean -> :boolean
  - Date -> :date
  - NaiveDateTime -> :datetime
  - Time -> :time
  - nil -> :blank

  ## Options

  - `:style` - A map of style properties
  - `:formula` - A formula string (sets type to :formula)

  ## Examples

      iex> Celex.Cell.new("Hello")
      %Celex.Cell{value: "Hello", type: :string, style: nil, formula: nil}

      iex> Celex.Cell.new(42, style: %{bold: true})
      %Celex.Cell{value: 42, type: :number, style: %{bold: true}, formula: nil}
  """
  @spec new(value(), keyword()) :: t()
  def new(value, opts \\ []) do
    style = Keyword.get(opts, :style)
    formula = Keyword.get(opts, :formula)

    type =
      cond do
        formula != nil -> :formula
        true -> infer_type(value)
      end

    %__MODULE__{
      value: value,
      type: type,
      style: style,
      formula: formula
    }
  end

  @doc """
  Creates a formula cell.

  If the formula starts with '=', it is stripped.

  ## Examples

      iex> Celex.Cell.formula("=SUM(A1:A10)")
      %Celex.Cell{value: nil, type: :formula, style: nil, formula: "SUM(A1:A10)"}

      iex> Celex.Cell.formula("AVERAGE(B1:B10)")
      %Celex.Cell{value: nil, type: :formula, style: nil, formula: "AVERAGE(B1:B10)"}
  """
  @spec formula(String.t()) :: t()
  def formula(formula_str) when is_binary(formula_str) do
    # Strip leading '=' if present
    formula = String.trim_leading(formula_str, "=")

    %__MODULE__{
      value: nil,
      type: :formula,
      style: nil,
      formula: formula
    }
  end

  @doc """
  Returns true if the cell is blank (nil value and no formula).

  ## Examples

      iex> Celex.Cell.blank?(Celex.Cell.new(nil))
      true

      iex> Celex.Cell.blank?(Celex.Cell.new(""))
      false
  """
  @spec blank?(t()) :: boolean()
  def blank?(%__MODULE__{value: nil, formula: nil}), do: true
  def blank?(%__MODULE__{}), do: false

  @doc """
  Returns true if the cell contains a formula.

  ## Examples

      iex> Celex.Cell.formula?(Celex.Cell.new(nil, formula: "SUM(A1:A10)"))
      true

      iex> Celex.Cell.formula?(Celex.Cell.new("Hello"))
      false
  """
  @spec formula?(t()) :: boolean()
  def formula?(%__MODULE__{formula: nil}), do: false
  def formula?(%__MODULE__{formula: _}), do: true

  # Type inference helpers

  defp infer_type(nil), do: :blank
  defp infer_type(value) when is_binary(value), do: :string
  defp infer_type(value) when is_integer(value), do: :number
  defp infer_type(value) when is_float(value), do: :number
  defp infer_type(value) when is_boolean(value), do: :boolean
  defp infer_type(%Date{}), do: :date
  defp infer_type(%NaiveDateTime{}), do: :datetime
  defp infer_type(%Time{}), do: :time
end
