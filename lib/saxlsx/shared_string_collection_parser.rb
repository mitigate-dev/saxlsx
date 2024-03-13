# frozen_string_literal: true
module Saxlsx
  class SharedStringCollectionParser < Ox::Sax

    def self.parse(file_system, &block)
      shared_strings = file_system.shared_strings
      if shared_strings
        SaxParser.parse self.new(&block), shared_strings
      else
        []
      end
    end

    def initialize(&block)
      @block = block
      @extract = false
    end

    def start_element(name)
      case name
      when :si then @current_string = nil
      when :t then @extract = true
      end
    end

    def end_element(name)
      case name
      when :si then @block.call(@current_string)
      when :t then @extract = false
      end
    end

    def text(value)
      if @extract
        if @current_string
          @current_string << CGI.unescapeHTML(value)
        else
          @current_string = CGI.unescapeHTML(value)
        end
      end
    end
  end
end
