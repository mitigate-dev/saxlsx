module Saxlsx
  class Sheet

    attr_reader :name, :state

    def initialize(name, state, index, file_system, workbook)
      @name = name
      @state = state
      @index = index
      @file_system = file_system
      @workbook = workbook
    end

    def rows
      @rows ||= RowsCollection.new(@index, @file_system, @workbook)
    end

    def to_csv(path)
      FileUtils.mkpath path unless Dir.exists? path
      File.open("#{path}/#{name}.csv", 'w') do |f|
        rows.each do |row|
          f.puts row.map{|c| "\"#{c}\""}.join(',')
        end
      end
    end

  end
end