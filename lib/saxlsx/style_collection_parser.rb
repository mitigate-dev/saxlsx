module Saxlsx
  class StyleCollectionParser < Ox::Sax
    def self.parse(file_system, &block)
      SaxParser.parse self.new(&block), file_system.styles
    end

    def initialize(&block)
      @block = block
      @cell_styles = false
    end

    def start_element(name)
      case name
      when :cellXfs
        @cell_styles = true
      when :xf
        @num_fmt_id = nil
      end
    end

    def end_element(name)
      case name
      when :cellXfs
        @cell_styles = false
      when :xf
        if @cell_styles
          @block.call @num_fmt_id
          @num_fmt_id = nil
        end
      end
    end

    def attr(name, value)
      if name == :numFmtId
        @num_fmt_id = value.to_i
      end
    end
  end
end