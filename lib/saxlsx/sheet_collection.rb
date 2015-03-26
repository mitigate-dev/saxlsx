module Saxlsx
  class SheetCollection

    include Enumerable

    def initialize(file_system, workbook)
      @file_system = file_system
      @workbook = workbook
    end

    def each(&block)
      SheetCollectionParser.parse @file_system, @workbook, &block
    end

  end
end