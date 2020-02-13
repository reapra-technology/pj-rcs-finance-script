class NetsuiteConsolidationReport
  include Util

  attr_accessor :input, :sheet_name,
    :sheet_reader, :sheet_writer, :sheet_props

  def initialize input, sheet_name
    self.input = input
    self.sheet_name = sheet_name
  end

  def prerequisite
    self.sheet_reader ||= Sheets::SheetReader.new input, sheet_name
    self.sheet_writer ||= Sheets::SheetWriter.new input, sheet_name
    self.sheet_props ||= Sheets::SheetProps.new get_sheet_read

    sheet_props.run

    self
  end

  def run
    # sheet_writer.read_config sheet_props.segments
    #
    # sheet_writer.do_something
    #
    # sheet_writer.write_output
  end

  def get_sheet_read
    sheet_reader.workbook.sheet(sheet_name)
  end

  def get_sheet_write
    sheet_reader.workbook.sheet(result_sheet_name)
  end

  def clone_entities
    left = get_coor_roo sheet_props.left.join
    right = get_coor_roo sheet_props.right.join

    x = left.first
    while x < right.first  do
      y = left.last
      while y < right.last do
       data = sheet_reader.read [x,y]
       sheet_writer.write data, coor_roo2_coor_rubyxl([x,y])
       y += 1
      end
      x +=1
    end

    sheet_writer.write_output
  end

  private
    def result_sheet_name
      "consol" + "_" + sheet_name
    end
end
