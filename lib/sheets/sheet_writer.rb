require 'rubyXL'
require 'rubyXL/convenience_methods'

module Sheet
  class SheetWriter
    attr_accessor :workbook, :filename, :worksheet

    def initialize filename = ''
      @filename = filename

      read_file
      add_new_console_worksheet
    end

    def read_file
      @workbook = RubyXL::Parser.parse filename
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

    private
      def get_coor x = 'A1'
        # A1 == [0, 0]
        return *RubyXL::Reference.ref2ind(x)
      end
  end
end
