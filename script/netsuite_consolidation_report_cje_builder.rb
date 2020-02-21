class NetsuiteConsolidationReportCjeBuilder
  include Util

  attr_accessor :input, :output, :bs, :pl, :cje_sheet, :verbose,
    :start_row_index, :start_col_index, :end_row_index

  def initialize input, output, bs, pl, cje_sheet = nil, verbose = false
    self.input = input
    self.output = output
    self.bs = bs
    self.pl = pl
    self.cje_sheet = cje_sheet
    self.verbose = verbose

    self.start_row_index = 4 # row 5
    self.end_row_index = 30 # row SUM
    self.start_col_index = 8 # column I

    # balancesheet_fill_color = "9dc3e6"
    # bs_subacc_fill_color = "c5e0b4"
    # pl_subacc_fill_color = "f8cbad"
  end

  def build_cje_sheet
    # TODO

    subacc_bs = get_list_subaccount_bs
    subacc_bs_length = subacc_bs.length

    subacc_pl = get_list_subaccount_pl
    subacc_pl_length = subacc_pl.length

    not_formula = false
    is_formula = true
    subacc_idx = start_col_index

    # write Balance Sheet + P/L and merged
    sheet_writer.write result_sheet_name, "BALANCE SHEET", [start_row_index, start_col_index - 4], not_formula
    sheet_writer.write result_sheet_name, "P/L", [start_row_index, start_col_index - 2], not_formula
    sheet_writer.merge_cells result_sheet_name,
      start_row_index , start_col_index - 4 , start_row_index , start_col_index - 3
    sheet_writer.merge_cells result_sheet_name,
      start_row_index , start_col_index - 2 , start_row_index , start_col_index - 1
    # formatting
    bs = sheet_writer.sheet_data_at result_sheet_name, start_row_index, start_col_index - 4
    pl = sheet_writer.sheet_data_at result_sheet_name, start_row_index, start_col_index - 2
    [bs,pl].each {|e|
      e.change_horizontal_alignment('center')
      e.change_vertical_alignment('center')
      e.change_border(:top, 'thin')
      e.change_border(:bottom, 'thin')
      e.change_border(:top, 'thin')
      e.change_border(:bottom, 'thin')
      e.change_border(:left, 'thin')
      e.change_border(:right, 'thin')
      e.change_font_bold(true)
      e.change_fill('9dc3e6')
    }

    # expand row height
    sheet_writer.workbook[result_sheet_name].change_row_height(start_row_index, 88)

    # populate subaccount
    [get_list_subaccount_bs, get_list_subaccount_pl].flatten.compact.each_with_index do |subacc, i|
      subacc_idx = start_col_index + i
      sheet_writer.write result_sheet_name, subacc, [start_row_index, subacc_idx], not_formula
      # formatting
      e = sheet_writer.sheet_data_at result_sheet_name, start_row_index, subacc_idx
      e.change_horizontal_alignment('center')
      e.change_vertical_alignment('center')
      e.change_border(:top, 'thin')
      e.change_border(:bottom, 'thin')
      e.change_border(:left, 'thin')
      e.change_border(:right, 'thin')
      e.change_text_wrap(true)
      i < subacc_bs_length ? e.change_fill('c5e0b4') : e.change_fill('f8cbad')
    end

    # populate check column
    sheet_writer.write result_sheet_name, "check", [start_row_index, subacc_idx + 1], not_formula
    # formatting
    e = sheet_writer.sheet_data_at result_sheet_name, start_row_index, subacc_idx + 1
    e.change_horizontal_alignment('center')
    e.change_vertical_alignment('center')
    e.change_border(:top, 'thin')
    e.change_border(:bottom, 'thin')
    e.change_border(:left, 'thin')
    e.change_border(:right, 'thin')
    e.change_text_wrap(true)

    # insert currency next line
    sheet_writer.write result_sheet_name, "Dr (S$)", [start_row_index+1, start_col_index - 4], not_formula
    sheet_writer.write result_sheet_name, "Cr (S$)", [start_row_index+1, start_col_index - 3], not_formula
    sheet_writer.write result_sheet_name, "Dr (S$)", [start_row_index+1, start_col_index - 2], not_formula
    sheet_writer.write result_sheet_name, "Cr (S$)", [start_row_index+1, start_col_index - 1], not_formula

    # insert check formula
    start_formula_row_idx = start_row_index+2
    (start_formula_row_idx .. end_row_index).each do |i|
      data = "SUM(#{column_name(start_col_index+1)}#{i+1}:#{column_name(subacc_idx)}#{i+1})-(#{column_name(start_col_index - 4)}#{i+1}-#{column_name(start_col_index - 3)}#{i+1})-(#{column_name(start_col_index - 2)}#{i+1}-#{column_name(start_col_index - 1)}#{i+1})"
      sheet_writer.write result_sheet_name, data, [i, subacc_idx + 1], is_formula
    end

    # insert sum below
    sum_row_idx = end_row_index + 4
    sheet_writer.write result_sheet_name, "SUM", [sum_row_idx, 0], not_formula
    (start_col_index - 4 .. subacc_idx + 1).each do |i|
      data = "SUM(#{column_name(i+1)}#{start_formula_row_idx+1}:#{column_name(i+1)}#{sum_row_idx})"
      sheet_writer.write result_sheet_name, data, [sum_row_idx, i], is_formula
    end
  end

  def save_file
    puts "Begin writing: #{Time.now}"
    sheet_writer.write_output output
    puts "End writing: #{Time.now}"
  end

  def get_list_subaccount_bs
    get_bs_sheet.column(1).select{|x| match_group_segment?(x) || match_child_segment?(x) }
  end

  def get_list_subaccount_pl
    get_pl_sheet.column(1).select{|x| match_group_segment?(x) || match_child_segment?(x) }
  end

  private
    def sheet_reader
      @sheet_reader ||= Sheets::SheetReader.new input
    end

    def sheet_writer
      @sheet_writer ||= Sheets::SheetWriter.new input
    end

    def result_sheet_name
      cje_sheet || "CJE"
    end

    def get_bs_sheet
      @get_bs_sheet ||= sheet_reader.workbook.sheet(bs).clone
    end

    def get_pl_sheet
      @get_pl_sheet ||= sheet_reader.workbook.sheet(pl).clone
    end

    def sheet_props_bs
      @sheet_props_bs ||= Sheets::SheetProps.new(get_bs_sheet)
    end

    def sheet_props_pl
      @sheet_props_pl ||= Sheets::SheetProps.new(get_pl_sheet)
    end
end
