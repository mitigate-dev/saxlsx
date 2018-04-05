# frozen_string_literal: true
module Saxlsx
  class StyleCollectionParser < Ox::Sax
    def self.parse(file_system, &block)
      SaxParser.parse self.new(&block), file_system.styles
    end

    def initialize(&block)
      @block = block
      @cell_styles = false
      @custom_num_fmts = {}
    end

    def start_element(name)
      case name
      when :cellXfs
        @cell_styles = true
      when :xf
        @num_fmt_id = nil
      when :numFmt
        @num_fmt_id = nil
        @num_fmt_code = nil
      end
    end

    def end_element(name)
      case name
      when :cellXfs
        @cell_styles = false
      when :xf
        if @cell_styles
          custom_num_fmt_code = @custom_num_fmts[@num_fmt_id]
          if custom_num_fmt_code
            @block.call custom_num_fmt_code
          else
            @block.call @num_fmt_id.to_i
          end
        end
      when :numFmt
        @custom_num_fmts[@num_fmt_id] = @num_fmt_code
      end
    end

    def attr(name, value)
      case name
      when :numFmtId
        @num_fmt_id = value.to_i
      when :formatCode
        @num_fmt_code = value
      end
    end
  end
end
