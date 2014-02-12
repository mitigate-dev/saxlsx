module Saxlsx
  class SheetCollection

    include Enumerable

    def initialize(file_system, shared_strings)
      @file_system = file_system
      @shared_strings = shared_strings
    end

    def each(&block)
      SheetCollectionParser.parse @file_system, @shared_strings, &block
    end

  end
end