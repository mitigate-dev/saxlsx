module Saxlsx
  class SheetCollectionParser < Ox::Sax

    CurrentSheet = Struct.new :index, :name

    def self.parse(file_system, shared_strings, &block)
      SaxParser.parse self.new(file_system, shared_strings, &block), file_system.workbook
    end

    def initialize(file_system, shared_strings, &block)
      @file_system = file_system
      @shared_strings = shared_strings
      @block = block
      @index = -1
    end

    def start_element(name)
      @current_sheet = CurrentSheet.new(@index += 1) if name == :sheet
    end

    def end_element(name)
      if name == :sheet
        @block.call Sheet.new(@current_sheet.name, @current_sheet.index, @file_system, @shared_strings)
        @current_sheet = nil
      end
    end

    def attr(name, value)
      @current_sheet.name = value if @current_sheet && name == :name
    end

  end
end