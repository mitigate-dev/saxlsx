# frozen_string_literal: true
module Saxlsx
  class ColumnNameGenerator
    FIRST = 'A'
    LAST  = 'Z'

    def self.next_to(previous)
      char = previous ? previous[-1] : nil
      if char.nil?
        FIRST
      elsif char < LAST
        previous[0..-2] + char.next
      else
        next_to(previous[0..-2]) + FIRST
      end
    end
  end
end
