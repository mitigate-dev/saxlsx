module Saxlsx
  class RowsCollection

    include Enumerable

    def initialize(index, file_system, shared_strings)
      @index = index
      @file_system = file_system
      @shared_strings = shared_strings
      @sheet = file_system.sheet(index)
    end

    def each(&block)
      RowsCollectionParser.parse @index, @sheet, @shared_strings, &block
    end

    def count
      @count ||= @sheet.match(/<dimension ref="[^:]+:[A-Z]+(\d+)"/)[1].to_i
    end

    alias :size :count

    def [](value)
      to_a[value]
    end
  end
end
