defmodule Celex.BenchmarkTest do
  use ExUnit.Case, async: false

  @moduletag :benchmark
  # 5 minute timeout for benchmarks
  @moduletag timeout: 300_000

  # Number of rows for benchmark tests
  @small_rows 10_000
  @medium_rows 100_000
  @large_rows 500_000

  setup do
    path = Path.join(System.tmp_dir!(), "benchmark_#{:erlang.unique_integer([:positive])}.xlsx")
    on_exit(fn -> File.rm(path) end)
    {:ok, path: path}
  end

  describe "memory benchmarks" do
    @tag :benchmark
    test "streaming write maintains constant memory for #{@small_rows} rows", %{path: path} do
      {memory_before, memory_after, duration_ms} = benchmark_write(path, @small_rows)

      # Memory should not grow significantly (allow 50MB headroom for GC variance)
      memory_growth_mb = (memory_after - memory_before) / 1_048_576

      IO.puts("\n[Benchmark] Write #{@small_rows} rows:")
      IO.puts("  Duration: #{duration_ms}ms")
      IO.puts("  Memory before: #{Float.round(memory_before / 1_048_576, 2)}MB")
      IO.puts("  Memory after: #{Float.round(memory_after / 1_048_576, 2)}MB")
      IO.puts("  Memory growth: #{Float.round(memory_growth_mb, 2)}MB")
      IO.puts("  Rows/second: #{Float.round(@small_rows / (duration_ms / 1000), 0)}")

      assert File.exists?(path)
      assert memory_growth_mb < 50, "Memory grew by #{memory_growth_mb}MB, expected < 50MB"
    end

    @tag :benchmark
    test "streaming read maintains constant memory for #{@small_rows} rows", %{path: path} do
      # First create the file
      create_test_file(path, @small_rows)

      {memory_before, memory_after, duration_ms, row_count} = benchmark_read(path)

      memory_growth_mb = (memory_after - memory_before) / 1_048_576

      IO.puts("\n[Benchmark] Read #{@small_rows} rows:")
      IO.puts("  Duration: #{duration_ms}ms")
      IO.puts("  Memory before: #{Float.round(memory_before / 1_048_576, 2)}MB")
      IO.puts("  Memory after: #{Float.round(memory_after / 1_048_576, 2)}MB")
      IO.puts("  Memory growth: #{Float.round(memory_growth_mb, 2)}MB")
      IO.puts("  Rows/second: #{Float.round(row_count / (duration_ms / 1000), 0)}")

      assert row_count == @small_rows
      assert memory_growth_mb < 50, "Memory grew by #{memory_growth_mb}MB, expected < 50MB"
    end

    @tag :benchmark
    @tag :slow
    test "streaming write handles #{@medium_rows} rows efficiently", %{path: path} do
      {memory_before, memory_after, duration_ms} = benchmark_write(path, @medium_rows)

      memory_growth_mb = (memory_after - memory_before) / 1_048_576

      IO.puts("\n[Benchmark] Write #{@medium_rows} rows:")
      IO.puts("  Duration: #{duration_ms}ms (#{Float.round(duration_ms / 1000, 2)}s)")
      IO.puts("  Memory growth: #{Float.round(memory_growth_mb, 2)}MB")
      IO.puts("  Rows/second: #{Float.round(@medium_rows / (duration_ms / 1000), 0)}")

      assert File.exists?(path)
      # Should complete in reasonable time
      assert duration_ms < 60_000, "Write took #{duration_ms}ms, expected < 60s"
    end

    @tag :benchmark
    @tag :slow
    test "streaming read handles #{@medium_rows} rows efficiently", %{path: path} do
      create_test_file(path, @medium_rows)

      {memory_before, memory_after, duration_ms, row_count} = benchmark_read(path)

      memory_growth_mb = (memory_after - memory_before) / 1_048_576

      IO.puts("\n[Benchmark] Read #{@medium_rows} rows:")
      IO.puts("  Duration: #{duration_ms}ms (#{Float.round(duration_ms / 1000, 2)}s)")
      IO.puts("  Memory growth: #{Float.round(memory_growth_mb, 2)}MB")
      IO.puts("  Rows/second: #{Float.round(row_count / (duration_ms / 1000), 0)}")

      assert row_count == @medium_rows
    end

    @tag :benchmark
    @tag :slow
    test "round-trip preserves data integrity for large files", %{path: path} do
      row_count = @small_rows

      # Write with known data pattern
      Celex.stream_write(path, fn writer ->
        writer
        |> Celex.StreamWriter.start_sheet("Data")
        |> Celex.StreamWriter.write_row(["ID", "Name", "Value", "Active"])
        |> Celex.StreamWriter.write_rows(
          Stream.map(1..row_count, fn i ->
            [i, "Item #{i}", i * 1.5, rem(i, 2) == 0]
          end)
        )
      end)

      # Read and verify
      {:ok, reader} = Celex.StreamReader.open(path)

      rows =
        reader
        |> Celex.StreamReader.stream_rows("Data")
        |> Enum.to_list()

      Celex.StreamReader.close(reader)

      # Verify header
      assert Enum.at(rows, 0) == ["ID", "Name", "Value", "Active"]

      # Verify first data row
      assert Enum.at(rows, 1) == [1, "Item 1", 1.5, false]

      # Verify last data row
      last_row = Enum.at(rows, row_count)
      assert Enum.at(last_row, 0) == row_count
      assert Enum.at(last_row, 1) == "Item #{row_count}"
      assert Enum.at(last_row, 2) == row_count * 1.5
      assert Enum.at(last_row, 3) == (rem(row_count, 2) == 0)

      # Verify total count (header + data rows)
      assert length(rows) == row_count + 1
    end

    @tag :benchmark
    @tag :stress
    test "handles #{@large_rows} rows without crashing", %{path: path} do
      IO.puts("\n[Stress Test] Writing #{@large_rows} rows...")

      # Use same measurement approach as other benchmarks for consistency
      {memory_before, memory_after, write_time} = benchmark_write(path, @large_rows)
      memory_growth_mb = (memory_after - memory_before) / 1_048_576

      assert File.exists?(path)

      {:ok, %{size: size}} = File.stat(path)
      size_mb = size / 1_048_576

      IO.puts("  Write time: #{Float.round(write_time / 1000, 2)}s")
      IO.puts("  File size: #{Float.round(size_mb, 2)}MB")
      IO.puts("  Memory growth: #{Float.round(memory_growth_mb, 2)}MB")
      IO.puts("  Rows/second: #{Float.round(@large_rows / (write_time / 1000), 0)}")

      # Read back and count rows
      IO.puts("  Reading back...")
      {_, _, read_time, row_count} = benchmark_read_by_sheet(path, "Data")

      IO.puts("  Read time: #{Float.round(read_time / 1000, 2)}s")
      IO.puts("  Rows/second: #{Float.round(row_count / (read_time / 1000), 0)}")

      assert row_count == @large_rows
    end
  end

  describe "file size benchmarks" do
    @tag :benchmark
    test "file size scales linearly with row count", %{path: path} do
      sizes =
        for rows <- [1000, 5000, 10000] do
          file_path = "#{path}_#{rows}"
          on_exit(fn -> File.rm(file_path) end)

          create_test_file(file_path, rows)
          {:ok, %{size: size}} = File.stat(file_path)

          {rows, size}
        end

      IO.puts("\n[Benchmark] File size scaling:")

      for {rows, size} <- sizes do
        IO.puts(
          "  #{rows} rows: #{Float.round(size / 1024, 2)}KB (#{Float.round(size / rows, 2)} bytes/row)"
        )
      end

      # Verify roughly linear scaling (within 2x variance)
      [{r1, s1}, {r2, s2}, {r3, s3}] = sizes
      bytes_per_row_1 = s1 / r1
      bytes_per_row_2 = s2 / r2
      bytes_per_row_3 = s3 / r3

      # All should be within 2x of each other
      avg = (bytes_per_row_1 + bytes_per_row_2 + bytes_per_row_3) / 3
      assert abs(bytes_per_row_1 - avg) / avg < 0.5
      assert abs(bytes_per_row_2 - avg) / avg < 0.5
      assert abs(bytes_per_row_3 - avg) / avg < 0.5
    end
  end

  # Helper functions

  defp benchmark_write(path, row_count) do
    :erlang.garbage_collect()
    memory_before = :erlang.memory(:total)

    start_time = System.monotonic_time(:millisecond)

    Celex.stream_write(path, fn writer ->
      writer
      |> Celex.StreamWriter.start_sheet("Data")
      |> Celex.StreamWriter.write_rows(
        Stream.map(1..row_count, fn i ->
          [i, "Name #{i}", i * 0.5, rem(i, 2) == 0]
        end)
      )
    end)

    duration_ms = System.monotonic_time(:millisecond) - start_time

    :erlang.garbage_collect()
    memory_after = :erlang.memory(:total)

    {memory_before, memory_after, duration_ms}
  end

  defp benchmark_read(path) do
    benchmark_read_by_sheet(path, "Data")
  end

  defp benchmark_read_by_sheet(path, sheet_name) do
    :erlang.garbage_collect()
    memory_before = :erlang.memory(:total)

    start_time = System.monotonic_time(:millisecond)

    row_count =
      path
      |> Celex.stream_sheet(sheet_name)
      |> Enum.count()

    duration_ms = System.monotonic_time(:millisecond) - start_time

    :erlang.garbage_collect()
    memory_after = :erlang.memory(:total)

    {memory_before, memory_after, duration_ms, row_count}
  end

  defp create_test_file(path, row_count) do
    Celex.stream_write(path, fn writer ->
      writer
      |> Celex.StreamWriter.start_sheet("Data")
      |> Celex.StreamWriter.write_rows(
        Stream.map(1..row_count, fn i ->
          [i, "Name #{i}", i * 0.5, rem(i, 2) == 0]
        end)
      )
    end)
  end
end
