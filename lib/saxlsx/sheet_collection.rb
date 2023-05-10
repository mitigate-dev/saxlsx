# frozen_string_literal: true
module Saxlsx
  class SheetCollection

    include Enumerable

    def initialize(file_system, workbook)
      @file_system = file_system
      @workbook = workbook
    end

    def each(&block)
      if block
        SheetCollectionParser.parse @file_system, @workbook, &block
      else
        to_enum
      end
    end

  end
end
