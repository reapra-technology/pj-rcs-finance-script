class NetsuiteConsolidationReportCjeBuilder
  include Util

  attr_accessor :input, :bs, :pl, :cje_sheet, :verbose

  def initialize input, bs, pl, cje_sheet = nil, verbose = false
    self.input = input
    self.bs = bs
    self.pl = pl
    self.cje_sheet = cje_sheet
    self.verbose = verbose
  end

  def build_cje_sheet
    # TODO
  end

  def get_list_subaccount_bs
    get_bs_sheet.column(1).select{|x| (x =~ /^\d\d\d\-\d\d\d\-\d\d\d/) || (x =~ /^\d\d\d\-\d\d\d\-\d\d\s\-/) }
  end

  def get_list_subaccount_pl
    get_pl_sheet.column(1).select{|x| (x =~ /^\d\d\d\-\d\d\d\-\d\d\d/) || (x =~ /^\d\d\d\-\d\d\d\-\d\d\s\-/) }
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
