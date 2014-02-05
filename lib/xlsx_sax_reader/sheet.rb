module XlsxSaxReader
  class Sheet

    attr_reader :name

    def initialize(name, index, file_system, shared_strings)
      @name = name
      @index = index
      @file_system = file_system
      @shared_strings = shared_strings
    end

    def rows
      @rows ||= RowsCollection.new(@index, @file_system, @shared_strings)
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