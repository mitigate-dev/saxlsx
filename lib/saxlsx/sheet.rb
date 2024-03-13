# frozen_string_literal: true
module Saxlsx
  class Sheet

    attr_reader :name

    def initialize(name, index, file_system, workbook)
      @name = name
      @index = index
      @file_system = file_system
      @workbook = workbook
    end

    def rows
      @rows ||= RowsCollection.new(@index, @file_system, @workbook)
    end

  end
end
