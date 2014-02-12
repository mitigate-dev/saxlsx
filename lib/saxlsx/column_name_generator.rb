module Saxlsx
  class ColumnNameGenerator

    FIRST = 'A'
    LAST = 'Z'

    def self.next_to(previous)
      cache previous do
        parts = (previous || '').chars.to_a
        char = parts.pop

        if char.nil? || char.empty?
          FIRST
        elsif char < LAST
          parts << (char.ord + 1).chr
          parts.join
        else
          "#{next_to(parts.join)}A"
        end
      end
    end

    private

    def self.cache(key, &block)
      @cache ||= {}

      return @cache[key] if @cache.has_key? key

      @cache[key] = block.call
    end

  end
end