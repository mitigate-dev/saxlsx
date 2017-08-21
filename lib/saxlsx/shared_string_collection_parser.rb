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
    end

    def start_element(name)
      @current_string = '' if name == :si
    end

    def end_element(name)
      if name == :si
        @block.call @current_string
        @current_string = nil
      end
    end

    def text(value)
      @current_string = CGI.unescapeHTML(value) if @current_string
    end

  end
end
