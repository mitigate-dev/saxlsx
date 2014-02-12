module Saxlsx
  class RowsCollection

    include Enumerable

    def initialize(index, file_system, shared_strings)
      @index = index
      @file_system = file_system
      @shared_strings = shared_strings
    end

    def each(&block)
      RowsCollectionParser.parse @index, @file_system, @shared_strings, &block
    end

    def [](value)
      to_a[value]
    end
  end
end