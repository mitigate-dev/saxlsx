# frozen_string_literal: true
module Saxlsx
  class RowsCollectionCountParser < Ox::Sax
    def self.count(data, &block)
      parser = new
      catch :abort do
        SaxParser.parse parser, data
      end
      parser.count
    end

    attr_reader :count

    def initialize
      @count = 0
    end

    def start_element(name)
      @current_element = name
      if name == :row
        @count += 1
      end
    end

    def attr(name, value)
      if @current_element == :dimension
        if name == :ref && value
          matches = value.match(/[^:]+:[A-Z]*(\d+)/)
          if matches
            @count = matches[1].to_i
            throw :abort
          end
        end
      end
    end
  end
end
