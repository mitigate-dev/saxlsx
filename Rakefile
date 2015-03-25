require "bundler/gem_tasks"
require "rspec/core/rake_task"

Bundler::GemHelper.install_tasks
RSpec::Core::RakeTask.new(:spec)
task :default => :spec


task :bench do
  require 'benchmark'
  require 'axlsx'
  require 'saxlsx'
  require 'rubyXL'
  require 'simple_xlsx_reader'
  require 'creek'

  path = "tmp/bench.xlsx"
  unless File.exists?(path)
    puts "* Generating #{path}"
    FileUtils.mkdir_p File.dirname(path)
    Axlsx::Package.new do |p|
      money_style = p.workbook.styles.add_style(
        num_fmt: 5, format_code: "€0.000"
      )
      p.workbook.add_worksheet(:name => "Pie Chart") do |sheet|
        10000.times do
          sheet.add_row(
            [Date.today, Time.now, 1000, 3.14, "Long" * 100],
            types: [:date, :time, :integer, :float, :string],
            style: [nil, nil, nil, money_style, nil]
          )
        end
      end
      p.use_shared_strings = true
      p.serialize(path)
    end
  end

  Benchmark.benchmark('', 20) do |x|
    x.report "creek" do
      w = Creek::Book.new path
      w.sheets.each do |s|
        s.rows.each do |r|
          r.values.inspect
        end
      end
    end
    x.report "rubyXL" do
      w = RubyXL::Parser.parse path
      w.worksheets.each do |s|
        s.each do |r|
          r.cells.map(&:value).inspect
        end
      end
    end
    x.report "saxlsx" do
      Saxlsx::Workbook.open path do |w|
        w.sheets.each do |s|
          s.rows.each do |r|
            r.to_a.inspect
          end
        end
      end
    end
    x.report "simple_xlsx_reader" do
      w = SimpleXlsxReader.open path
      w.sheets.each do |s|
        s.rows.each do |r|
          r.to_a.inspect
        end
      end
    end
  end
end