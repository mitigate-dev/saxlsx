# frozen_string_literal: true
module Saxlsx
  class SaxParser

    def self.parse(handler, xml)
      Ox.sax_parse(handler, xml)
    ensure
      xml.rewind
    end

  end
end
