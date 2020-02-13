module Sheets
  class SheetWriter
    include Util

    attr_accessor :workbook, :filename, :sheet_name, :worksheet

    def initialize filename = '', sheet_name
      self.filename = filename
      self.sheet_name = sheet_name
      read_file
      add_new_console_worksheet
    end

    def read_file
      @workbook = RubyXL::Parser.parse filename
      @workbook.calc_pr.full_calc_on_load = true
    end

    def add_new_console_worksheet
      @worksheet = @workbook.add_worksheet "consol_" + sheet_name
    end

    def write data, coor, is_formula = false
      is_formula ?
        @worksheet.add_cell(coor.first, coor.last, '', data) :
        @worksheet.add_cell(coor.first, coor.last, data)
    end

    def write_output
      # Write
      # worksheet.add_cell(0, 0, 'A1')      # Sets cell A1 to string "A1"
      # worksheet.add_cell(0, 1, '', 'A1')  # Sets formula in the cell B1 to '=A1'
      # worksheet.add_cell(0,2,'', 'PI()/4')
      # worksheet.add_cell(1,2,'', 'PI()*4')
      # worksheet.add_cell(2,2,'', 'C1*C2')
      puts "Begin writing: #{Time.now}"
      @workbook.write("./output/output.xlsx")
      puts "End writing: #{Time.now}"
    end
  end
end
