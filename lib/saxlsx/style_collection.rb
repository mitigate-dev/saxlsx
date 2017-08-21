# frozen_string_literal: true
module Saxlsx
  class StyleCollection

    include Enumerable

    def initialize(file_system)
      @file_system = file_system
    end

    def each(&block)
      StyleCollectionParser.parse @file_system, &block
    end

  end
end
