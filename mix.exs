defmodule Celex.MixProject do
  use Mix.Project

  @version "0.1.0"
  @source_url "https://github.com/qlamu/celex"

  def project do
    [
      app: :celex,
      version: @version,
      elixir: "~> 1.14",
      start_permanent: Mix.env() == :prod,
      deps: deps(),
      description: description(),
      package: package(),
      docs: docs(),
      name: "Celex",
      source_url: @source_url
    ]
  end

  def application do
    [
      extra_applications: [:logger, :xmerl]
    ]
  end

  defp deps do
    [
      {:ex_doc, "~> 0.31", only: :dev, runtime: false}
    ]
  end

  defp description do
    """
    A robust, stream-based Elixir library for reading and writing XLSX files.
    Handles large files (millions of rows) with constant memory usage.
    """
  end

  defp package do
    [
      licenses: ["MIT"],
      links: %{"GitHub" => @source_url}
    ]
  end

  defp docs do
    [
      main: "readme",
      extras: ["README.md"]
    ]
  end
end
