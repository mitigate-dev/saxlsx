# frozen_string_literal: true
module Saxlsx
  class ColumnNameGenerator
    FIRST = 'A'
    LAST  = 'Z'

    def self.next_to(previous)
      if previous
        previous.succ
      else
        FIRST
      end
    end
  end
end
