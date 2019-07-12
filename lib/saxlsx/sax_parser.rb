# frozen_string_literal: true
module Saxlsx
  class SaxParser

    def self.parse(handler, xml)
      Ox.sax_parse(handler, xml, skip: :skip_return)
    ensure
      xml.rewind
    end

  end
end
