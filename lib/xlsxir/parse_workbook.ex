defmodule Xlsxir.ParseWorkbook do
  @moduledoc """
  Holds the SAX event instructions for parsing workbook data via `Xlsxir.SaxParser.parse/2`
  """

  defstruct name: "", index: 1, tid: nil

  @doc """
  Sax event utilized by `Xlsxir.SaxParser.parse/2`. Takes a pattern and the current state of a struct and recursivly parses the
  styles XML file, ultimately saving each parsed style type to the ETS process.

  ## Parameters

  - pattern - the XML pattern of the event to match upon
  - state - the state of the `%Xlsxir.ParseWorkbook{}` struct

  ## Example
  Recursively sends workbook sheet information generated from parsing the `xl/workbook.xml` file to ETS process. The data can ultimately
  be retreived from the ETS table (i.e. `:ets.lookup(tid, 1)` would return something "Sheet1"
  """
  def sax_event_handler(:startDocument, _state) do
    %__MODULE__{tid: GenServer.call(Xlsxir.StateManager, :new_table)}
  end

  def sax_event_handler({:startElement, _, 'sheet', _, xml_attr}, state) do
    {_, _, _, _, name} = Enum.find(xml_attr, fn {_, a, _, _, _} -> a == 'name' end)
    %{state | name: name |> to_string}
  end

  def sax_event_handler(
        {:endElement, _, 'sheet', _},
        %__MODULE__{name: name, index: index, tid: tid} = state
      ) do
    :ets.insert(tid, {index, name})
    %{state | index: index + 1}
  end

  def sax_event_handler(_, state), do: state
end
