# frozen_string_literal: true
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
      if block
        RowsCollectionParser.parse @index, @sheet, @workbook, &block
      else
        to_enum
      end
    end

    def count
      @count ||= RowsCollectionCountParser.count @sheet
    end

    alias :size :count

    def [](value)
      to_a[value]
    end
  end
end
