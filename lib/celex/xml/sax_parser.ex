defmodule Celex.XML.SaxParser do
  @moduledoc """
  SAX-based XML parser for streaming XLSX worksheet data.

  Uses Erlang's :xmerl_sax_parser for event-based parsing, yielding
  rows one at a time to achieve O(1) memory usage regardless of file size.
  """

  alias Celex.Utils.CellReference

  @doc """
  Parses worksheet XML and streams rows to the caller process.

  Spawns a parser process that sends `{:row, row_data}` messages
  for each row, and `{:done, :ok}` when complete.
  """
  @spec stream_rows(binary(), [String.t()], pid()) :: :ok
  def stream_rows(xml, shared_strings, caller_pid) do
    state = %{
      caller: caller_pid,
      shared_strings: shared_strings,
      in_row: false,
      in_cell: false,
      in_value: false,
      in_inline_string: false,
      in_inline_text: false,
      current_row: [],
      current_cell_ref: nil,
      current_cell_type: nil,
      current_value: "",
      row_index: 0
    }

    event_fun = &handle_event/3

    try do
      :xmerl_sax_parser.stream(
        xml,
        [{:event_fun, event_fun}, {:event_state, state}]
      )

      send(caller_pid, {:done, :ok})
    catch
      :throw, {:halt, _state} ->
        send(caller_pid, {:done, :ok})

      kind, reason ->
        send(caller_pid, {:error, {kind, reason}})
    end

    :ok
  end

  # SAX Event Handlers

  defp handle_event({:startElement, _uri, local_name, _qname, attrs}, _location, state) do
    name = to_string(local_name)
    handle_start_element(name, attrs, state)
  end

  defp handle_event({:endElement, _uri, local_name, _qname}, _location, state) do
    name = to_string(local_name)
    handle_end_element(name, state)
  end

  defp handle_event({:characters, chars}, _location, state) do
    handle_characters(to_string(chars), state)
  end

  defp handle_event(_event, _location, state) do
    state
  end

  # Start element handlers

  defp handle_start_element("row", attrs, state) do
    _row_num = get_attr(attrs, "r")
    %{state | in_row: true, current_row: [], row_index: state.row_index + 1}
  end

  defp handle_start_element("c", attrs, state) when state.in_row do
    ref = get_attr(attrs, "r")
    type = get_attr(attrs, "t")
    %{state | in_cell: true, current_cell_ref: ref, current_cell_type: type, current_value: ""}
  end

  defp handle_start_element("v", _attrs, state) when state.in_cell do
    %{state | in_value: true, current_value: ""}
  end

  defp handle_start_element("is", _attrs, state) when state.in_cell do
    %{state | in_inline_string: true}
  end

  defp handle_start_element("t", _attrs, state) when state.in_inline_string do
    %{state | in_inline_text: true, current_value: ""}
  end

  defp handle_start_element(_name, _attrs, state), do: state

  # End element handlers

  defp handle_end_element("row", state) do
    # Build the complete row and send it
    row = build_row(state.current_row)
    send(state.caller, {:row, row})
    %{state | in_row: false, current_row: []}
  end

  defp handle_end_element("c", state) when state.in_cell do
    # Add cell to current row
    cell = {state.current_cell_ref, state.current_cell_type, state.current_value}

    %{
      state
      | in_cell: false,
        current_row: [cell | state.current_row],
        current_cell_ref: nil,
        current_cell_type: nil,
        current_value: ""
    }
  end

  defp handle_end_element("v", state) when state.in_value do
    %{state | in_value: false}
  end

  defp handle_end_element("is", state) when state.in_inline_string do
    %{state | in_inline_string: false}
  end

  defp handle_end_element("t", state) when state.in_inline_text do
    %{state | in_inline_text: false}
  end

  defp handle_end_element(_name, state), do: state

  # Character data handler

  defp handle_characters(chars, %{in_value: true} = state) do
    %{state | current_value: state.current_value <> chars}
  end

  defp handle_characters(chars, %{in_inline_text: true} = state) do
    %{state | current_value: state.current_value <> chars}
  end

  defp handle_characters(_chars, state), do: state

  # Helper functions

  defp get_attr(attrs, name) do
    name_charlist = to_charlist(name)

    Enum.find_value(attrs, fn
      {_uri, _prefix, local_name, value} when local_name == name_charlist ->
        to_string(value)

      _ ->
        nil
    end)
  end

  defp build_row(cells) do
    # Cells are in reverse order, and we need to handle gaps
    cells
    |> Enum.reverse()
    |> Enum.map(&parse_cell/1)
    |> fill_row_gaps()
  end

  defp parse_cell({ref, type, value}) do
    col_index = parse_column_index(ref)
    cell_value = convert_value(type, value)
    {col_index, cell_value}
  end

  defp parse_column_index(nil), do: nil

  defp parse_column_index(ref) do
    case Regex.run(~r/^([A-Z]+)/, ref) do
      [_, col] -> CellReference.column_to_index(col)
      nil -> nil
    end
  end

  defp convert_value("s", value) do
    # Shared string - will be resolved by caller
    {:shared_string, parse_int(value)}
  end

  defp convert_value("b", value) do
    value == "1"
  end

  defp convert_value("inlineStr", value) do
    unescape_xml(value)
  end

  defp convert_value(_type, "") do
    nil
  end

  defp convert_value(_type, value) do
    case Float.parse(value) do
      {float, ""} ->
        if trunc(float) == float, do: trunc(float), else: float

      _ ->
        unescape_xml(value)
    end
  end

  defp parse_int(str) do
    case Integer.parse(str) do
      {n, _} -> n
      :error -> 0
    end
  end

  defp fill_row_gaps([]), do: []

  defp fill_row_gaps(cells) do
    cells = Enum.filter(cells, fn {col, _} -> col != nil end)

    if cells == [] do
      []
    else
      max_col = cells |> Enum.map(fn {col, _} -> col end) |> Enum.max()
      cells_map = Map.new(cells)

      for col <- 1..max_col do
        Map.get(cells_map, col)
      end
    end
  end

  defp unescape_xml(text) when is_binary(text) do
    text
    |> String.replace("&amp;", "&")
    |> String.replace("&lt;", "<")
    |> String.replace("&gt;", ">")
    |> String.replace("&quot;", "\"")
    |> String.replace("&apos;", "'")
  end

  defp unescape_xml(text), do: to_string(text)
end
