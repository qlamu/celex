defmodule Celex.Utils.DateUtils do
  @moduledoc """
  Utilities for converting between Elixir dates/times and Excel serial date numbers.

  Excel stores dates as floating-point numbers where:
  - The integer part represents days since December 30, 1899
  - The fractional part represents time as a fraction of 24 hours

  ## Excel Date System Notes

  Excel has a famous bug: it incorrectly treats 1900 as a leap year, meaning
  February 29, 1900 (serial number 60) doesn't actually exist. This module
  accounts for this bug to ensure compatibility with Excel files.

  Serial number reference points:
  - 1 = January 1, 1900
  - 60 = (fake) February 29, 1900
  - 61 = March 1, 1900
  - 36526 = January 1, 2000

  ## Examples

      iex> Celex.Utils.DateUtils.date_to_serial(~D[2000-01-01])
      36526

      iex> Celex.Utils.DateUtils.serial_to_date(36526)
      ~D[2000-01-01]
  """

  # Excel's epoch is December 31, 1899 (serial 0)
  # Serial 1 = January 1, 1900
  @excel_epoch Date.new!(1899, 12, 31)

  # Serial number 60 is the non-existent February 29, 1900
  # Dates after this need adjustment due to Excel's leap year bug
  @leap_year_bug_threshold 60

  @seconds_per_day 86_400

  @doc """
  Converts an Elixir Date to an Excel serial number.

  ## Examples

      iex> Celex.Utils.DateUtils.date_to_serial(~D[1900-01-01])
      1

      iex> Celex.Utils.DateUtils.date_to_serial(~D[2000-01-01])
      36526
  """
  @spec date_to_serial(Date.t()) :: integer()
  def date_to_serial(%Date{} = date) do
    days = Date.diff(date, @excel_epoch)

    # Account for Excel's 1900 leap year bug
    # For dates on or after March 1, 1900, we add 1 to account for
    # the non-existent Feb 29, 1900 that Excel thinks exists
    if days >= @leap_year_bug_threshold do
      days + 1
    else
      days
    end
  end

  @doc """
  Converts an Excel serial number to an Elixir Date.

  ## Examples

      iex> Celex.Utils.DateUtils.serial_to_date(1)
      ~D[1900-01-01]

      iex> Celex.Utils.DateUtils.serial_to_date(36526)
      ~D[2000-01-01]
  """
  @spec serial_to_date(number()) :: Date.t()
  def serial_to_date(serial) when is_number(serial) do
    days = trunc(serial)

    # Adjust for Excel's leap year bug
    adjusted_days =
      if days > @leap_year_bug_threshold do
        days - 1
      else
        days
      end

    Date.add(@excel_epoch, adjusted_days)
  end

  @doc """
  Converts an Elixir NaiveDateTime to an Excel serial number (with fractional time).

  ## Examples

      iex> Celex.Utils.DateUtils.datetime_to_serial(~N[2000-01-01 12:00:00])
      36526.5
  """
  @spec datetime_to_serial(NaiveDateTime.t()) :: float()
  def datetime_to_serial(%NaiveDateTime{} = datetime) do
    date = NaiveDateTime.to_date(datetime)
    time = NaiveDateTime.to_time(datetime)

    date_serial = date_to_serial(date)
    time_fraction = time_to_serial(time)

    date_serial + time_fraction
  end

  @doc """
  Converts an Excel serial number to an Elixir NaiveDateTime.

  ## Examples

      iex> Celex.Utils.DateUtils.serial_to_datetime(36526.5)
      ~N[2000-01-01 12:00:00]
  """
  @spec serial_to_datetime(number()) :: NaiveDateTime.t()
  def serial_to_datetime(serial) when is_number(serial) do
    date = serial_to_date(serial)
    time = serial_to_time(serial)

    NaiveDateTime.new!(date, time)
  end

  @doc """
  Converts an Elixir Time to an Excel serial fraction (0.0 to < 1.0).

  ## Examples

      iex> Celex.Utils.DateUtils.time_to_serial(~T[12:00:00])
      0.5

      iex> Celex.Utils.DateUtils.time_to_serial(~T[06:00:00])
      0.25
  """
  @spec time_to_serial(Time.t()) :: float()
  def time_to_serial(%Time{} = time) do
    seconds = Time.diff(time, ~T[00:00:00], :second)
    seconds / @seconds_per_day
  end

  @doc """
  Converts an Excel serial fraction to an Elixir Time.

  Takes either a fractional part (0.0-1.0) or extracts the fractional
  part from a full serial number.

  ## Examples

      iex> Celex.Utils.DateUtils.serial_to_time(0.5)
      ~T[12:00:00]

      iex> Celex.Utils.DateUtils.serial_to_time(36526.5)
      ~T[12:00:00]
  """
  @spec serial_to_time(number()) :: Time.t()
  def serial_to_time(serial) when is_number(serial) do
    # Get just the fractional part
    fraction = serial - trunc(serial)

    total_seconds = round(fraction * @seconds_per_day)

    hours = div(total_seconds, 3600)
    remaining = rem(total_seconds, 3600)
    minutes = div(remaining, 60)
    seconds = rem(remaining, 60)

    Time.new!(hours, minutes, seconds)
  end
end
