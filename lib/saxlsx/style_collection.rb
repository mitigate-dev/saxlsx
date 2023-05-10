# frozen_string_literal: true
module Saxlsx
  class StyleCollection

    include Enumerable

    def initialize(file_system)
      @file_system = file_system
    end

    def each(&block)
      if block
        StyleCollectionParser.parse @file_system, &block
      else
        to_enum
      end
    end

  end
end
