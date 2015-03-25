module Saxlsx
  class FileSystem

    def self.open(filename)
      begin
        file_system = self.new(filename)
        yield file_system
      ensure
        file_system.close
      end
    end

    def initialize(filename)
      @zip = Zip::File.open filename
    end

    def close
      @zip.close
    end

    def workbook
      @zip.read('xl/workbook.xml').match(/<sheets>.*<\/sheets>/).to_s
    end

    def shared_strings
      @zip.read('xl/sharedStrings.xml')
    end

    def styles
      @zip.read('xl/styles.xml')
    end

    def sheet(i)
      f = @zip.glob('xl/worksheets/sheet*.xml').sort[i]
      @zip.read(f)
    end

  end
end