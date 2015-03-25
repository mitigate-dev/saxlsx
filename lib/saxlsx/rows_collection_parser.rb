module Saxlsx
  class RowsCollectionParser < Ox::Sax
    NUM_FORMATS = {
      0  => :string,         # General
      1  => :fixnum,         # 0
      2  => :float,          # 0.00
      3  => :fixnum,         # #,##0
      4  => :float,          # #,##0.00
      5  => :unsupported,    # $#,##0_);($#,##0)
      6  => :unsupported,    # $#,##0_);[Red]($#,##0)
      7  => :unsupported,    # $#,##0.00_);($#,##0.00)
      8  => :unsupported,    # $#,##0.00_);[Red]($#,##0.00)
      9  => :percentage,     # 0%
      10 => :percentage,     # 0.00%
      11 => :bignum,         # 0.00E+00
      12 => :unsupported,    # # ?/?
      13 => :unsupported,    # # ??/??
      14 => :date,           # mm-dd-yy
      15 => :date,           # d-mmm-yy
      16 => :date,           # d-mmm
      17 => :date,           # mmm-yy
      18 => :time,           # h:mm AM/PM
      19 => :time,           # h:mm:ss AM/PM
      20 => :time,           # h:mm
      21 => :time,           # h:mm:ss
      22 => :date_time,      # m/d/yy h:mm
      37 => :unsupported,    # #,##0 ;(#,##0)
      38 => :unsupported,    # #,##0 ;[Red](#,##0)
      39 => :unsupported,    # #,##0.00;(#,##0.00)
      40 => :unsupported,    # #,##0.00;[Red](#,##0.00)
      45 => :time,           # mm:ss
      46 => :time,           # [h]:mm:ss
      47 => :time,           # mmss.0
      48 => :bignum,         # ##0.0E+0
      49 => :unsupported     # @
    }

    DATE_SYSTEM_1900 = DateTime.new(1899, 12, 30)

    def self.parse(index, data, workbook, &block)
      SaxParser.parse self.new(workbook, &block), data
    end

    def initialize(workbook, &block)
      @shared_strings = workbook.shared_strings
      @number_formats = workbook.number_formats
      @block = block
    end

    def start_element(name)
      @current_element = name

      if name == :row
        @current_row = []
        @next_column = 'A'
      end

      if name == :c
        @current_type = nil
        @current_number_format = nil
      end
    end

    def end_element(name)
      if name == :row
        @block.call @current_row
        @current_row = nil
      end
    end

    def attr(name, value)
      if @current_element == :c
        case name
        when :t
          @current_type = value
        when :r
          @current_column = value.gsub(/\d/, '')
        when :s
          @current_number_format = NUM_FORMATS[@number_formats[value.to_i]]
        end
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
        case @current_number_format
        when :date
          DATE_SYSTEM_1900 + text.to_i
        when :date_time
          # Round time to seconds
          date = DATE_SYSTEM_1900 + (text.to_f * 86400).round.fdiv(86400)
          DateTime.new(date.year, date.month, date.day, date.hour, date.minute, date.second)
        when :fixnum
          text.to_i
        when :float
          text.to_f
        when :bignum
          BigDecimal.new(text)
        when :percentage
          text.to_f / 100
        else
          # Auto convert numbers
          if text =~ /\A-?\d+(\.\d+(?:e[+-]\d+)?)?\Z/i
            $1 ? text.to_f : text.to_i
          else
            CGI.unescapeHTML(text)
          end
        end
      end
    end

  end
end