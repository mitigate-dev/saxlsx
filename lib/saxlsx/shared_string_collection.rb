module Saxlsx
  class SharedStringCollection

    include Enumerable

    def initialize(file_system)
      @file_system = file_system
    end

    def each(&block)
      SharedStringCollectionParser.parse @file_system, &block
    end

  end
end