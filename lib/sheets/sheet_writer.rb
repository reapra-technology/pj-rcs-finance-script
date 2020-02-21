module Sheets
  class SheetWriter
    include Util

    attr_accessor :workbook, :filename

    def initialize filename = ''
      self.filename = filename
      read_file
    end

    def read_file
      @workbook = RubyXL::Parser.parse filename
      @workbook.calc_pr.full_calc_on_load = true
    end

    def write sheet_name, data, coor, is_formula = false
      sheet = worksheet sheet_name
      is_formula ?
        sheet.add_cell(coor.first, coor.last, '', data) :
        sheet.add_cell(coor.first, coor.last, data)
    end

    def write_output output = nil
      @workbook.write(output || "./output/output.xlsx")
    end

    private
      def worksheet sheet_name
        @workbook[sheet_name] || @workbook.add_worksheet(sheet_name)
      end
  end
end
