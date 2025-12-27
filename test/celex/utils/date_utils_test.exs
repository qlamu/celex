defmodule Celex.Utils.DateUtilsTest do
  use ExUnit.Case, async: true

  alias Celex.Utils.DateUtils

  describe "date_to_serial/1" do
    test "converts epoch date" do
      # Excel serial number 1 = January 1, 1900
      assert DateUtils.date_to_serial(~D[1900-01-01]) == 1
    end

    test "converts early dates" do
      assert DateUtils.date_to_serial(~D[1900-01-02]) == 2
      assert DateUtils.date_to_serial(~D[1900-02-28]) == 59
    end

    test "accounts for Excel's 1900 leap year bug" do
      # Excel incorrectly thinks 1900 is a leap year
      # March 1, 1900 should be serial 61 (skipping the fake Feb 29)
      assert DateUtils.date_to_serial(~D[1900-03-01]) == 61
    end

    test "converts modern dates" do
      # Well-known reference: January 1, 2000 = 36526
      assert DateUtils.date_to_serial(~D[2000-01-01]) == 36526
      # December 31, 2024 = 36526 + Date.diff(~D[2024-12-31], ~D[2000-01-01]) = 36526 + 9131
      assert DateUtils.date_to_serial(~D[2024-12-31]) == 45657
    end

    test "converts various dates" do
      assert DateUtils.date_to_serial(~D[1999-12-31]) == 36525
      assert DateUtils.date_to_serial(~D[2020-01-01]) == 43831
      assert DateUtils.date_to_serial(~D[2025-06-15]) == 45823
    end
  end

  describe "serial_to_date/1" do
    test "converts serial 1 to epoch" do
      assert DateUtils.serial_to_date(1) == ~D[1900-01-01]
    end

    test "converts early serial numbers" do
      assert DateUtils.serial_to_date(2) == ~D[1900-01-02]
      assert DateUtils.serial_to_date(59) == ~D[1900-02-28]
    end

    test "handles Excel's 1900 leap year bug" do
      # Serial 60 is the fake Feb 29, 1900 - we skip it and return March 1
      # Actually we should handle this more gracefully - return March 1 for 60
      assert DateUtils.serial_to_date(61) == ~D[1900-03-01]
    end

    test "converts modern serial numbers" do
      assert DateUtils.serial_to_date(36526) == ~D[2000-01-01]
      assert DateUtils.serial_to_date(45657) == ~D[2024-12-31]
    end

    test "handles zero (returns epoch date)" do
      # Serial 0 = December 31, 1899 (day before Jan 1, 1900)
      assert DateUtils.serial_to_date(0) == ~D[1899-12-31]
    end
  end

  describe "datetime_to_serial/1" do
    test "converts datetime at midnight" do
      dt = ~N[2000-01-01 00:00:00]
      assert DateUtils.datetime_to_serial(dt) == 36526.0
    end

    test "converts datetime at noon" do
      dt = ~N[2000-01-01 12:00:00]
      assert_in_delta DateUtils.datetime_to_serial(dt), 36526.5, 0.0001
    end

    test "converts datetime with time" do
      dt = ~N[2000-01-01 06:00:00]
      assert_in_delta DateUtils.datetime_to_serial(dt), 36526.25, 0.0001

      dt = ~N[2000-01-01 18:00:00]
      assert_in_delta DateUtils.datetime_to_serial(dt), 36526.75, 0.0001
    end

    test "converts datetime with minutes and seconds" do
      # 1 hour = 1/24 = 0.041666...
      # 1 minute = 1/1440 = 0.000694...
      # 1 second = 1/86400 = 0.000011574...
      dt = ~N[2000-01-01 01:30:30]
      expected = 36526.0 + 1.0 / 24 + 30.0 / 1440 + 30.0 / 86400
      assert_in_delta DateUtils.datetime_to_serial(dt), expected, 0.00001
    end
  end

  describe "serial_to_datetime/1" do
    test "converts integer serial to datetime at midnight" do
      assert DateUtils.serial_to_datetime(36526) == ~N[2000-01-01 00:00:00]
      assert DateUtils.serial_to_datetime(36526.0) == ~N[2000-01-01 00:00:00]
    end

    test "converts fractional serial to datetime with time" do
      result = DateUtils.serial_to_datetime(36526.5)
      assert result.hour == 12
      assert result.minute == 0
      assert result.second == 0
    end

    test "converts various fractional times" do
      result = DateUtils.serial_to_datetime(36526.25)
      assert result.hour == 6
      assert result.minute == 0

      result = DateUtils.serial_to_datetime(36526.75)
      assert result.hour == 18
      assert result.minute == 0
    end
  end

  describe "time_to_serial/1" do
    test "converts midnight" do
      assert DateUtils.time_to_serial(~T[00:00:00]) == 0.0
    end

    test "converts noon" do
      assert_in_delta DateUtils.time_to_serial(~T[12:00:00]), 0.5, 0.0001
    end

    test "converts various times" do
      assert_in_delta DateUtils.time_to_serial(~T[06:00:00]), 0.25, 0.0001
      assert_in_delta DateUtils.time_to_serial(~T[18:00:00]), 0.75, 0.0001
      assert_in_delta DateUtils.time_to_serial(~T[23:59:59]), 0.99999, 0.0001
    end
  end

  describe "serial_to_time/1" do
    test "converts 0 to midnight" do
      assert DateUtils.serial_to_time(0.0) == ~T[00:00:00]
    end

    test "converts 0.5 to noon" do
      assert DateUtils.serial_to_time(0.5) == ~T[12:00:00]
    end

    test "converts various fractional values" do
      assert DateUtils.serial_to_time(0.25) == ~T[06:00:00]
      assert DateUtils.serial_to_time(0.75) == ~T[18:00:00]
    end

    test "extracts time from full serial number" do
      # 36526.5 = Jan 1, 2000 at noon - should extract just the time
      assert DateUtils.serial_to_time(36526.5) == ~T[12:00:00]
    end
  end

  describe "roundtrip" do
    test "date roundtrip" do
      dates = [
        ~D[1900-01-01],
        ~D[1900-03-01],
        ~D[2000-01-01],
        ~D[2024-06-15],
        ~D[2099-12-31]
      ]

      for date <- dates do
        serial = DateUtils.date_to_serial(date)
        assert DateUtils.serial_to_date(serial) == date
      end
    end

    test "datetime roundtrip (within second precision)" do
      datetimes = [
        ~N[2000-01-01 00:00:00],
        ~N[2000-01-01 12:00:00],
        ~N[2024-06-15 08:30:00],
        ~N[2024-12-31 23:59:59]
      ]

      for dt <- datetimes do
        serial = DateUtils.datetime_to_serial(dt)
        result = DateUtils.serial_to_datetime(serial)
        assert NaiveDateTime.diff(result, dt, :second) == 0
      end
    end
  end
end
