module Saxlsx
  class RowsCollectionParser < Ox::Sax

    def self.parse(index, file_system, shared_strings, &block)
      SaxParser.parse self.new(shared_strings, &block), file_system.sheets[index]
    end

    def initialize(shared_strings, &block)
      @shared_strings = shared_strings
      @block = block
    end

    def start_element(name)
      @current_element = name

      if name == :row
        @current_row = []
        @next_column = 'A'
      end

      @current_type = nil if name == :c
    end

    def end_element(name)
      if name == :row
        @block.call @current_row
        @current_row = nil
      end
    end

    def attr(name, value)
      if @current_element == :c
        @current_type = value if name == :t
        @current_column = value.gsub(/\d/, '') if name == :r
      end
    end

    def text(value)
      if @current_row && @current_element == :v
        while @next_column != @current_column
          @current_row << nil
          @next_column = ColumnNameGenerator.next_to(@next_column)
        end
        @current_row << value_of(value)
        @next_column = ColumnNameGenerator.next_to(@next_column)
      end
    end

    private

    def value_of(text)
      case @current_type
        when 's'
          @shared_strings[text.to_i]
        when 'b'
          BooleanParser.parse text
        when 'n'
          text.to_f
        else
          text
      end

    end

  end
end