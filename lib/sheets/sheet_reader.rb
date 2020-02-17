module Sheets
  class SheetReader
    attr_accessor :workbook, :worksheet

    def initialize filename, sheet_name = nil
      read_file(filename)
      read_sheet(sheet_name) if sheet_name
    end

    def read_file filename
      @workbook = Roo::Spreadsheet.open filename, extension: :xlsx
    end

    def read_sheet sheet_name
      return if @workbook.nil? || sheet_name.nil?

      @worksheet = workbook.sheet(sheet_name)
    end
  end
end
