# https://github.com/weshatheleopard/rubyXL
require 'rubyXL'
xl = workbook = RubyXL::Parser.parse("src.xlsx") 
 
def print_sheet(sheet)
  sheet.each { |row|
     row && row.cells.each { |cell|
       val = cell && cell.value
       puts(val)
     }
  }
end

print_sheet(xl[0])