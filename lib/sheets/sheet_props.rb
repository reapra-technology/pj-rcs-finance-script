module Sheets
  class SheetProps
    include Util

    attr_accessor :sheet, :segments,
      :left, :left_col, :left_row,
      :right, :right_col, :right_row,
      :z_elim_index, :z_elim_col,
      :locked_full_coor,
      :cje_sheet, :cje_sheet_left_col, :cje_sheet_left_row,
      :cje_sheet_right_col, :cje_sheet_right_row,
      :cje_locked_full_coor

    PL_GROUP_SEGMENTS = ["Sales", "Purchases", "Unrealized Matching Gain/Loss"]
    BS_GROUP_SEGMENTS = ["VAT on Sales BD", "Cumulative Translation Adjustment-Elimination", "Retained Earnings", "Cumulative Translation Adjustment"]

    def initialize sheet, cje_sheet = nil
      self.sheet = sheet
      self.cje_sheet = cje_sheet
    end

    def run
      get_coordinates
      get_cje_coordinates if cje_sheet
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

        self.segments = res.map{|x| Range.new(col.find_index(x.first) + 1, col.find_index(x.last).nil? ? col.find_index(x.first) + 1 : col.find_index(x.last) + 1)}
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

        self.locked_full_coor = [
          left.map{|e| "$#{e.to_s.upcase}" }.join,
          right.map{|e| "$#{e.to_s.upcase}" }.join
        ].join(":")
      end

      def get_cje_coordinates
        cje_sheet.each_with_index {|r, i|
          row_index = i + 1
          if self.cje_sheet_left_row.nil? && r.include?("BALANCE SHEET") && r.include?("P/L")
            self.cje_sheet_left_row = row_index
          end

          if self.cje_sheet_right_row.nil? && r.first && r.first.to_s.upcase == "SUM"
            self.cje_sheet_right_row = row_index
          end
        }

        self.cje_sheet_right_col = column_name(cje_sheet.row(self.cje_sheet_left_row).length-1)
        self.cje_sheet_left_col = column_name(
          cje_sheet.row(cje_sheet_left_row).find_index{|x| x =~ /^\d\d\d\-\d\d\d\-/} + 1
        )
        self.cje_locked_full_coor = "$#{cje_sheet_left_col}$#{cje_sheet_left_row}:$#{cje_sheet_right_col}$#{cje_sheet_right_row}"
      end

      def match_group_segment? str
        !!(str =~ /^\d\d\d\-\d\d\d\s\-/ || (BS_GROUP_SEGMENTS + PL_GROUP_SEGMENTS).include?(str.to_s))
      end

      def match_child_segment? str
        !!((str =~ /^\d\d\d\-\d\d\d\-\d\d\d/) || (str =~ /^\d\d\d\-\d\d\d\-\d\d\s\-/))
      end
  end
end
