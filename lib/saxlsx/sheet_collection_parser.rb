# frozen_string_literal: true
module Saxlsx
  class SheetCollectionParser < Ox::Sax

    CurrentSheet = Struct.new :index, :name, :state

    def self.parse(file_system, workbook, &block)
      SaxParser.parse(
        self.new(file_system, workbook, &block),
        file_system.workbook
      )
    end

    def initialize(file_system, workbook, &block)
      @file_system = file_system
      @workbook = workbook
      @block = block
      @index = -1
      @workbook_pr = false
    end

    def start_element(name)
      case name
      when :sheet
        @current_sheet = CurrentSheet.new(@index += 1)
      when :workbookPr
        @workbook_pr = true
      end
    end

    def end_element(name)
      case name
      when :sheet
        @block.call Sheet.new(
          @current_sheet.name,
          @current_sheet.state,
          @current_sheet.index,
          @file_system,
          @workbook
        )
        @current_sheet = nil
      when :workbookPr
        @workbook_pr = false
      end
    end

    def attr(name, value)
      if @current_sheet
        if name == :name
          @current_sheet.name = value
        elsif name == :state
          @current_sheet.state = value
        end
      elsif @workbook_pr
        if name == :date1904 && value =~ /true|1/i
          @workbook.date1904 = true
        end
      end
    end

  end
end
