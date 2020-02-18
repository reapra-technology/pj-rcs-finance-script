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

    def write_output
      puts "Begin writing: #{Time.now}"
      @workbook.write("./output/output.xlsx")
      puts "End writing: #{Time.now}"
    end

    private

      def worksheet sheet_name
        @workbook["consol_" + sheet_name] || @workbook.add_worksheet("consol_" + sheet_name)
      end
  end
end
