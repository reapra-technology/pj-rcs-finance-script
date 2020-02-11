class SheetProps
  attr_accessor :sheet, :segments,
    :left, :left_col, :left_row,
    :right, :right_col, :right_row,
    :z_elim_index, :z_elim_col

  def initialize sheet
    self.sheet = sheet
  end

  def run
    get_coordinates
    get_segments
  end

  private
    def get_segments
      col = sheet.column("A")
      res = []

      col.each_with_index {|col_val, i|
        if match_group_segment? col_val
          res << [col_val, "Total - " + col_val]
        end
      }

      self.segments = res.map{|x| [col.find_index(x.first) + 1, col.find_index(x.last).nil? ? nil : col.find_index(x.last) + 1]}
    end

    def get_coordinates
      # assumption
      self.left_col = 'a'

      sheet.each_with_index {|r, i|
        row_index = i + 1
        if r[0] == 'Financial Row'
          self.left_row = row_index
          break
        end
      }

      self.right_row = sheet.column('a').size
      self.right_col = column_name sheet.row(self.left_row).find_index{|x| x == 'Total'} + 1

      self.left = [left_col, left_row]
      self.right = [right_col, right_row]

      self.z_elim_index = sheet.row(7).find_index{|x| x.start_with? 'z-Elim'} + 1
      self.z_elim_col = column_name z_elim_index
    end

    def excel_col_index(str)
      offset = 'A'.ord - 1
      str.upcase.chars.inject(0){ |x,c| x*26 + c.ord - offset }
    end

    def column_name(int)
      name = 'A'
      (int - 1).times { name.succ! }
      name
    end

    def match_group_segment? str
      !(str =~ /^\d\d\d\-\d\d\d\s\-/).nil?
    end

    def match_child_segment? str
      !(str =~ /^\d\d\d\-\d\d\d\-\d\d\d/).nil?
    end
end
