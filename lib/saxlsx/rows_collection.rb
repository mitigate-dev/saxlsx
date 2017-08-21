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
      RowsCollectionParser.parse @index, @sheet, @workbook, &block
    end

    def count
      unless defined?(@count)
        @count = 0
        begin
          @sheet.each_line('>') do |line|
            matches = line.match(/<dimension ref="[^:]+:[A-Z]*(\d+)"/)
            if matches
              @count = matches[1].to_i
              break if @count
            end
          end
        ensure
          @sheet.rewind
        end
      end
      @count
    end

    alias :size :count

    def [](value)
      to_a[value]
    end
  end
end
