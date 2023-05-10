# frozen_string_literal: true
module Saxlsx
  class RowsCollectionParser < Ox::Sax
    SECONDS_IN_DAY = 86400
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
      12 => :rational,       # # ?/?
      13 => :rational,       # # ??/??
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

    def self.parse(index, data, workbook, &block)
      SaxParser.parse self.new(workbook, &block), data
    end

    def initialize(workbook, &block)
      @base_date      = workbook.base_date
      @auto_format    = workbook.auto_format
      @shared_strings = workbook.shared_strings
      @number_formats = workbook.number_formats
      @block = block
    end

    def start_element(name)
      @current_element = name
      case name
      when :row
        @current_row = []
      when :c
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
          # fill missing columns with nil values
          nb_columns = extract_current_column(value) - 1
          while @current_row.size != nb_columns
            @current_row << nil
          end
        when :s
          @current_number_format = detect_format_type(value.to_i)
        end
      end
    end

    def text(value)
      if @current_row && (@current_element == :v || @current_element == :t)
        @current_row << value_of(value)
      end
    end

    private

    def value_of(text)
      case @current_type
      when 's'
        @shared_strings[text.to_i]
      when 'inlineStr'
        CGI.unescapeHTML(text)
      when 'b'
        BooleanParser.parse text
      else
        case @current_number_format
        when :date
          @base_date + Float(text)
        when :date_time
          # Round time to seconds
          date = @base_date + Rational((Float(text) * SECONDS_IN_DAY).round, SECONDS_IN_DAY)
          DateTime.new(date.year, date.month, date.day, date.hour, date.minute, date.second)
        when :fixnum
          Integer(text, 10)
        when :float, :percentage
          Float(text)
        when :rational
          Rational(text)
        when :bignum
          Float(text) # raises ArgumentError if text is not a number
          BigDecimal(text) # doesn't raise ArgumentError
        else
          if @current_type == 'n'
            Float(text)
          elsif @auto_format && text =~ /\A-?\d+(\.\d+(?:e[+-]\d+)?)?\Z/i
            # Auto convert numbers
            $1 ? Float(text) : Integer(text, 10)
          else
            CGI.unescapeHTML(text)
          end
        end
      end
    rescue ArgumentError
      CGI.unescapeHTML(text)
    end

    def detect_format_type(index)
      format = @number_formats[index]
      NUM_FORMATS[format] || detect_custom_format_type(format)
    end

    # This is the least deterministic part of reading xlsx files. Due to
    # custom styles, you can't know for sure when a date is a date other than
    # looking at its format and gessing. It's not impossible to guess right,
    # though.
    #
    # http://stackoverflow.com/questions/4948998/determining-if-an-xlsx-cell-is-date-formatted-for-excel-2007-spreadsheets
    def detect_custom_format_type(code)
      code = code.gsub(/\[[^\]]+\]/, '') # Strip meta - [...]
      if code =~ /0/
        :float
      elsif code =~ /[ymdhis]/i
        :date_time
      else
        :unsupported
      end
    end

    def char_index(byte)
      if byte >= 65 && byte <= 90
        byte - 64
      elsif byte >= 97 && byte <= 122
        byte - 96
      end
    end

    def extract_current_column(value)
      letter_num = 0

      value.each_byte do |b|
        if index = char_index(b)
          letter_num *= 26
          letter_num += index
        else
          break
        end
      end

      raise ArgumentError if letter_num == 0

      letter_num
    end
  end
end
