module Sheets
  class SheetReader
    attr_accessor :workbook, :filename

    def initialize filename = ''
      @filename = filename

      read_file
    end

    def read_file
      @workbook = Roo::Spreadsheet.open @filename, extension: :xlsx
    end
  end
end
