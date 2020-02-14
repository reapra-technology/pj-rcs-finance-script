class NetsuiteConsolidationReport
  include Util

  attr_accessor :input, :sheet_name

  def initialize input, sheet_name
    self.input = input
    self.sheet_name = sheet_name
  end

  def prerequisite
    sheet_props.run

    self
  end

  def run
    sheet_writer.write_output
  end

  def get_sheet_read
    sheet_reader.workbook.sheet(sheet_name)
  end

  def get_sheet_write
    sheet_reader.workbook.sheet(result_sheet_name)
  end

  def get_segments
    sheet_props.segments
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

  def clone_entities_with_ref_formulas
    left = get_coor_roo sheet_props.left.join
    right = get_coor_roo sheet_props.right.join

    # clone skeleton
    x = left.first
    is_formula = true
    while x < right.first  do
      y = left.last
      while y < right.last do
        if x == left.first || y == left.last
          data = "\'#{sheet_name}\'!#{get_ref(coor_roo2_coor_rubyxl([x,y]))}"
          sheet_writer.write data, coor_roo2_coor_rubyxl([x,y]), is_formula
        else
          apply_vlookup_formulas x, y
        end
        y += 1
      end
      x +=1
      puts "Row: #{x}"
    end
  end

  def apply_vlookup_formulas x,y
    get_segments.each {|r|
      if r.include? x
        data = "VLOOKUP(A#{x},\'#{sheet_name}\'!#{sheet_props.locked_full_coor},#{y},FALSE)"
        sheet_writer.write data, coor_roo2_coor_rubyxl([x,y]), true
      end
    }
  end

  private
    def sheet_reader
      @sheet_reader ||= Sheets::SheetReader.new input, sheet_name
    end

    def sheet_writer
      @sheet_writer ||= Sheets::SheetWriter.new input, sheet_name
    end

    def sheet_props
      @sheet_props ||= Sheets::SheetProps.new get_sheet_read
    end

    def get_segments
      @segments ||= sheet_props.segments
    end

    def result_sheet_name
      "consol" + "_" + sheet_name
    end
end
