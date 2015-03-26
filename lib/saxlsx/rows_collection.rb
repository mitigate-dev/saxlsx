module Saxlsx
  class RowsCollection

    include Enumerable

    def initialize(index, file_system, workbook)
      @index = index
      @file_system = file_system
      @workbook = workbook
      @sheet = file_system.sheet(index)
    end

    def each(&block)
      RowsCollectionParser.parse @index, @sheet, @workbook, &block
    end

    def count
      @count ||= @sheet.match(/<dimension ref="[^:]+:[A-Z]*(\d+)"/)[1].to_i
    end

    alias :size :count

    def [](value)
      to_a[value]
    end
  end
end
