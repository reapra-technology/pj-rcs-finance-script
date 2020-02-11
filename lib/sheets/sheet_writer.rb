module Sheets
  class SheetWriter
    include Util

    attr_accessor :workbook, :filename, :worksheet

    def initialize filename = ''
      @filename = filename

      read_file
      add_new_console_worksheet
    end

    def read_file
      @workbook = RubyXL::Parser.parse filename
      @workbook.calc_pr.full_calc_on_load = true
    end

    def add_new_console_worksheet
      @worksheet = @workbook.add_worksheet "consol_" + filename
    end

    # https://github.com/weshatheleopard/rubyXL
    def write
      wb = xlsx_package.workbook
      wb.add_worksheet(name: "Buttons") do |sheet|
        @buttons.each do |button|
          sheet.add_row [button.name, button.category, button.price]
        end
      end
    end

    def write_output
      # Write
      # worksheet.add_cell(0, 0, 'A1')      # Sets cell A1 to string "A1"
      # worksheet.add_cell(0, 1, '', 'A1')  # Sets formula in the cell B1 to '=A1'
      # worksheet.add_cell(0,2,'', 'PI()/4')
      # worksheet.add_cell(1,2,'', 'PI()*4')
      # worksheet.add_cell(2,2,'', 'C1*C2')
      @workbook.write("./output/#{filename}")
    end
  end
end
