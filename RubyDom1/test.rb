require "google_drive"

session = GoogleDrive::Session.from_config("config.json")
$spreadsheet = session.spreadsheet_by_key("1rwfAax11xRiqGEoKELU42pneJMoY-8fDmiIlVSXOcH8")
sheet = $spreadsheet.worksheets[0]
sheet1 = $spreadsheet.worksheets[1]

class Table
  include Enumerable
  attr_reader :sheet

  def initialize(sheet)
    @sheet = sheet
    @table = sheet.rows.reject do |row|
      row.all? { |cell| cell.nil? || cell.strip.empty? } ||
      row.any? { |cell| cell.to_s.strip.downcase.include?("total") || cell.to_s.strip.downcase.include?("subtotal") }
    end
  end

  def table
    @table
  end

  def row(row)
    @table[row]
  end

  def [](name)
    idx = @sheet.rows[0].index(name)
    raise "Nije pronadjena" unless idx
    Column.new(@sheet, idx)
  end

  def print_table(tabl)
    tabl.table.each do |row|
      puts row.inspect
    end
  end

  def each()
    @table.each { |row| row.each {|n| yield n}}
  end

  def +(table1)
    raise "Nisu isti hederi" unless table1.table.first == @table.first

    combo_sheet = $spreadsheet.worksheets[2]
    table = Table.new(combo_sheet)
    comb = (@table + table1.table).uniq
    comb.each_with_index do |row, idx|
      row.each_with_index { |cell, idx_c| table.sheet[idx + 1,idx_c + 1] = cell}
    end
    table.sheet.save
    table
  end

  def -(table1)
    raise "Nisu isti hederi" unless table1.table.first == @table.first

    combo_sheet = $spreadsheet.worksheets[3]
    table = Table.new(combo_sheet)

    comb = @table.reject do |row|
      table1.table.any? { |row_t| row == row_t }
    end
    comb.unshift(@table.first)
    comb.each_with_index do |row, idx|
      row.each_with_index { |cell, idx_c| table.sheet[idx + 1,idx_c + 1] = cell}
    end
    table.sheet.save
    table
  end

  def method_missing(name, *args)
    str = name.to_s.downcase.gsub(/\s+/, "")
    @table[0].each_with_index do |n, idx|
      head = n.to_s.downcase.gsub(/\s+/, "")
      if str == head
        return Column.new(@sheet, idx)
      end
    end
    super
  end

    class Column
        def initialize(sheet, idx)
          @sheet = sheet
          @table = sheet.rows
          @idx = idx
        end

        def [](row)
          @table[row + 1, @idx + 1]
        end

        def []=(row, value)
          @sheet[row + 1, @idx + 1] = value
          @sheet.save
        end

        def map(&block)
          @table.each_with_index do |row, row_index|
            next if row_index.zero?
            next if  row.all? { |cell| cell.nil? || cell.strip.empty? } || row[@idx].to_s.downcase.include?('total') || row[@idx].to_s.downcase.include?('subtotal')

            cell_value = row[@idx]
            new_value = yield(cell_value)
            @sheet[row_index + 1, @idx + 1] = new_value
          end
          @sheet.save
        end

        def select
          sel = []
          @table.drop(1).each do |row|
            sel << row[@idx] if yield(row[@idx])
          end
          sel
        end

        def reduce(ini)
          acc = ini
          @tabela.drop(1).each { |row|   acc = yield(acc, row[@idx])}
        end

        def each
          @table.each {|row| yield row[@idx]}
        end

        def method_missing(name, *args, &block)
          niz = []
          self.each {|k| niz << k}
          niz.each_with_index do |n , row_idx|
            if name.to_s == n
              return @table[row_idx]
            end
          end
            super
        end

        def sum
          @table.map { |row| row[@idx].to_i }.reduce(0, :+)
        end

        def avg
          sum / (@table.size - 1)
        end
    end
end



table = Table.new(sheet)
table1 = Table.new(sheet1)
# p table.row(1)
# p "--------------"
#  table.each {|k| p k}
# p "--------------"
# p table["Prva Kolona"][1]
# p "--------------"
# p table["Prva Kolona"].class
# p "--------------"
# table["Prva Kolona"][2]= 2551
# p "--------------"
# table.index.each {|k| p k}
# p "--------------"
# p table.prvaKolona.avg
# p "--------------"
# p table.index.rn10722
# p "--------------"
#  p table.prvaKolona.map { |cell| cell.to_i + 2 }
# p "--------------"
# p table.prvaKolona[2]
# p "--------------"
# p table.prvaKolona.select { |value| value.to_i > 10 }
# puts ""
# p "--------table---------"
# table.print_table(table)
# puts ""
# p "--------table1--------"
# table1.print_table(table1)
# puts ""
# p "--------Combo(+)------"
# tableCombP = table + table1;
# puts "done"
# p "--------Combo(-)------"
# tableCombM = table1 - table;
# puts "done"
