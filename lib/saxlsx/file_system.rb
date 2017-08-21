# frozen_string_literal: true
module Saxlsx
  class FileSystem
    IO_METHODS = [:tell, :seek, :read, :close]

    def self.open(filename)
      begin
        file_system = self.new(filename)
        yield file_system
      ensure
        file_system.close
      end
    end

    def initialize(filename)
      if IO_METHODS.map { |method| filename.respond_to?(method) }.all?
        @zip = Zip::File.open_buffer filename
        @io = true
      else
        @zip = Zip::File.open filename
      end
    end

    def close
      @zip.close unless @io
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
