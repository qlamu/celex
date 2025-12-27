defmodule Celex.Workbook do
  @moduledoc """
  Represents an Excel workbook containing one or more worksheets.

  A Workbook contains:
  - `sheets` - List of Sheet structs
  - `active_sheet` - The name of the active (visible) sheet
  - `metadata` - Document properties (creator, title, etc.)

  For streaming operations, use `Celex.StreamWriter` instead.

  ## Examples

      workbook = Workbook.new()
      |> Workbook.add_sheet("Data")
      |> Workbook.add_sheet("Summary")

      # Update a sheet
      workbook = Workbook.update_sheet(workbook, "Data", fn sheet ->
        sheet
        |> Sheet.put_cell("A1", "Hello")
        |> Sheet.put_cell("B1", "World")
      end)
  """

  alias Celex.Sheet

  @type t :: %__MODULE__{
          sheets: [Sheet.t()],
          active_sheet: String.t() | nil,
          metadata: %{atom() => any()}
        }

  defstruct sheets: [], active_sheet: nil, metadata: %{}

  @doc """
  Creates a new empty workbook.
  """
  @spec new() :: t()
  def new do
    %__MODULE__{}
  end

  @doc """
  Adds a sheet to the workbook.

  Can accept a sheet name (string) or a Sheet struct.
  The first sheet added becomes the active sheet.
  """
  @spec add_sheet(t(), String.t() | Sheet.t()) :: t()
  def add_sheet(%__MODULE__{sheets: sheets, active_sheet: active} = workbook, name)
      when is_binary(name) do
    sheet = Sheet.new(name, index: length(sheets))
    new_sheets = sheets ++ [sheet]
    new_active = active || name

    %{workbook | sheets: new_sheets, active_sheet: new_active}
  end

  def add_sheet(%__MODULE__{sheets: sheets, active_sheet: active} = workbook, %Sheet{} = sheet) do
    sheet_with_index = %{sheet | index: length(sheets)}
    new_sheets = sheets ++ [sheet_with_index]
    new_active = active || sheet.name

    %{workbook | sheets: new_sheets, active_sheet: new_active}
  end

  @doc """
  Gets a sheet by name or index.

  Index is 0-based.
  """
  @spec get_sheet(t(), String.t() | non_neg_integer()) :: Sheet.t() | nil
  def get_sheet(%__MODULE__{sheets: sheets}, name) when is_binary(name) do
    Enum.find(sheets, fn sheet -> sheet.name == name end)
  end

  def get_sheet(%__MODULE__{sheets: sheets}, index) when is_integer(index) do
    Enum.at(sheets, index)
  end

  @doc """
  Updates a sheet by name using the given function.

  The function receives the sheet and should return the updated sheet.
  """
  @spec update_sheet(t(), String.t(), (Sheet.t() -> Sheet.t())) :: t()
  def update_sheet(%__MODULE__{sheets: sheets} = workbook, name, fun) when is_function(fun, 1) do
    updated_sheets =
      Enum.map(sheets, fn sheet ->
        if sheet.name == name do
          fun.(sheet)
        else
          sheet
        end
      end)

    %{workbook | sheets: updated_sheets}
  end

  @doc """
  Returns list of sheet names in order.
  """
  @spec sheet_names(t()) :: [String.t()]
  def sheet_names(%__MODULE__{sheets: sheets}) do
    Enum.map(sheets, & &1.name)
  end

  @doc """
  Returns the number of sheets.
  """
  @spec sheet_count(t()) :: non_neg_integer()
  def sheet_count(%__MODULE__{sheets: sheets}) do
    length(sheets)
  end

  @doc """
  Sets the active (visible) sheet by name.
  """
  @spec set_active_sheet(t(), String.t()) :: t()
  def set_active_sheet(%__MODULE__{} = workbook, name) when is_binary(name) do
    %{workbook | active_sheet: name}
  end

  @doc """
  Removes a sheet by name.
  """
  @spec remove_sheet(t(), String.t()) :: t()
  def remove_sheet(%__MODULE__{sheets: sheets, active_sheet: active} = workbook, name) do
    new_sheets = Enum.reject(sheets, fn sheet -> sheet.name == name end)

    new_active =
      if active == name do
        case new_sheets do
          [first | _] -> first.name
          [] -> nil
        end
      else
        active
      end

    %{workbook | sheets: new_sheets, active_sheet: new_active}
  end

  @doc """
  Renames a sheet.
  """
  @spec rename_sheet(t(), String.t(), String.t()) :: t()
  def rename_sheet(
        %__MODULE__{sheets: sheets, active_sheet: active} = workbook,
        old_name,
        new_name
      ) do
    updated_sheets =
      Enum.map(sheets, fn sheet ->
        if sheet.name == old_name do
          %{sheet | name: new_name}
        else
          sheet
        end
      end)

    new_active =
      if active == old_name do
        new_name
      else
        active
      end

    %{workbook | sheets: updated_sheets, active_sheet: new_active}
  end

  @doc """
  Sets a metadata property.

  Common properties:
  - `:creator` - Document author
  - `:title` - Document title
  - `:subject` - Document subject
  - `:description` - Document description
  - `:created` - Creation timestamp
  - `:modified` - Last modified timestamp
  """
  @spec set_metadata(t(), atom(), any()) :: t()
  def set_metadata(%__MODULE__{metadata: metadata} = workbook, key, value) do
    %{workbook | metadata: Map.put(metadata, key, value)}
  end

  @doc """
  Gets a metadata property.
  """
  @spec get_metadata(t(), atom()) :: any()
  def get_metadata(%__MODULE__{metadata: metadata}, key) do
    Map.get(metadata, key)
  end
end
