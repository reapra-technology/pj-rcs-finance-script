module Util
  # [x, y]
  def get_coor x = 'A1'
    # A1 == [0, 0]
    # A5 == [4, 0]
    # AA5 == [4, 26]
    return RubyXL::Reference.ref2ind(x.upcase)
  end

  def get_coor_roo x = 'A1'
    return RubyXL::Reference.ref2ind(x.upcase).map{|e| e+1}
  end

  def get_ref idx = [0,0]
    # => A1
    return RubyXL::Reference.ind2ref(*idx)
  end

  def coor_roo2_coor_rubyxl coor = [0,0]
    # [-1, -1]
    coor.map{|e| e-1}
  end

  def excel_col_index str
    # A => 1
    # AA => 27
    offset = 'A'.ord - 1
    str.upcase.chars.inject(0){ |x,c| x*26 + c.ord - offset }
  end

  def column_name int
    # 27 => AA
    # 50 => AX
    name = 'A'
    (int - 1).times { name.succ! }
    name
  end

  module Roo
    module Excelx
      def read coor
        self.cell(*coor)
      end
    end
  end
end
