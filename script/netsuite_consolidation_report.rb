class NetsuiteConsolidationReport
  include Util

  attr_accessor :input, :sheet_name

  def initialize input, sheet_name
    self.input = input
    self.sheet_name = sheet_name
  end

  def prerequisite
    sheet_reader
    sheet_props.run

  end

  def run
    sheet_writer.read_config sheet_props.segments
    #
    sheet_writer.do_something
    #
    sheet_writer.write_output
  end

  private
    def sheet_reader
      @sheet_reader ||= Sheets::SheetReader.new input
    end

    def sheet_writer
      @sheet_writer ||= Sheets::SheetWriter.new input
    end

    def sheet_props
      @sheet_props ||= Sheets::SheetProps.new get_sheet
    end

    def result_sheet_name
      "consol" + "_" + sheet_name
    end

    def get_sheet
      sheet_reader.workbook.sheet(sheet_name)
    end
end
