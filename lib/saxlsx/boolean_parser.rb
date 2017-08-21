# frozen_string_literal: true
module Saxlsx
  class BooleanParser

    def self.parse(string)
      return true if string == true || string =~ (/(true|t|yes|y|1)$/i)
      return false if string == false || string.nil? || string =~ (/(false|f|no|n|0)$/i)
      raise ArgumentError.new("Invalid value for Boolean: \"#{string}\"")
    end

  end
end
