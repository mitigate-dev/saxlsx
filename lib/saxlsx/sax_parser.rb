module Saxlsx
  class SaxParser

    def self.parse(handler, xml)
      io = StringIO.new xml
      Ox.sax_parse(handler, io)
      io.close
    end

  end
end