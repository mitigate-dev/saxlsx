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
      @zip.get_input_stream('xl/workbook.xml')
    end

    def shared_strings
      file = @zip.glob('xl/shared[Ss]trings.xml').first
      @zip.get_input_stream(file) if file
    end

    def styles
      @zip.get_input_stream('xl/styles.xml')
    end

    def sheet(i)
      @zip.get_input_stream("xl/worksheets/sheet#{i+1}.xml")
    end

  end
end