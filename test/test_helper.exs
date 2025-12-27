# Exclude benchmark and stress tests by default
# Run with: mix test --include benchmark
# Run with: mix test --include stress
ExUnit.start(exclude: [:benchmark, :stress, :slow])
