require 'benchmark'
require 'axlsx'
require 'saxlsx'
require 'rubyXL'
require 'simple_xlsx_reader'
require 'creek'
require 'oxcelix'
require 'roo'
require 'dullard'

module Saxlsx
  class Benchmarks
    def run
      path = "tmp/bench_shared_strings.xlsx"
      generate path, true
      benchmark "Shared Strings", path

      path = "tmp/bench_inline_strings.xlsx"
      generate path, false
      benchmark "Inline Strings", path
    end

    private

    def generate(path, shared_strings)
      unless File.exists?(path)
        puts "* Generating #{path}"
        FileUtils.mkdir_p File.dirname(path)
        Axlsx::Package.new do |p|
          money_style = p.workbook.styles.add_style(
            num_fmt: 5, format_code: "€0.000"
          )
          p.workbook.add_worksheet(:name => "Sheet 1") do |sheet|
            10000.times do
              sheet.add_row(
                [Date.today, Time.now, 1000, 3.14, "Long" * 100],
                types: [:date, :time, :integer, :float, :string],
                style: [nil, nil, nil, money_style, nil]
              )
            end
          end
          p.use_shared_strings = shared_strings
          p.serialize(path)
        end
      end
    end

    def benchmark(title, path)
      puts
      puts title
      puts
      Benchmark.bmbm(20) do |x|
        x.report "creek" do
          run_creek(path)
        end
        x.report "dullard" do
          run_dullard(path)
        end
        x.report "oxcelix" do
          run_oxcelix(path)
        end
        x.report "roo" do
          run_roo(path)
        end
        x.report "rubyXL" do
          run_rubyxl(path)
        end
        x.report "saxlsx" do
          run_saxlsx(path)
        end
        x.report "simple_xlsx_reader" do
          run_simple_xlsx_reader(path)
        end
      end
    end

    def run_creek(path)
      w = Creek::Book.new path
      w.sheets.each do |s|
        s.rows.each do |r|
          r.values.inspect
        end
      end
    end

    def run_oxcelix(path)
      w = Oxcelix::Workbook.new(path)
      w.sheets.each do |s|
        s.to_ru.to_a.each do |r|
          r.inspect
        end
      end
    rescue
      puts "ERROR"
    end

    def run_rubyxl(path)
      w = RubyXL::Parser.parse path
      w.worksheets.each do |s|
        s.each do |r|
          r.cells.map(&:value).inspect
        end
      end
    end

    def run_saxlsx(path)
      Saxlsx::Workbook.open path do |w|
        w.sheets.each do |s|
          s.rows.each do |r|
            r.to_a.inspect
          end
        end
      end
    end

    def run_simple_xlsx_reader(path)
      w = SimpleXlsxReader.open path
      w.sheets.each do |s|
        s.rows.each do |r|
          r.to_a.inspect
        end
      end
    end

    def run_roo(path)
      w = Roo::Excelx.new path
      w.each_with_pagename do |_, s|
        s.each do |r|
          r.to_a.inspect
        end
      end
    end

    def run_dullard(path)
      w = Dullard::Workbook.new path
      w.sheets.each do |s|
        s.rows.each do |r|
          r.to_a.inspect
        end
      end
    end
  end
end
