# https://github.com/roo-rb/roo
require 'roo'

def sheet_to_arr(sheet_name)
  xl = Roo::Spreadsheet.open('src.xlsx')
  sheet = xl.sheet(sheet_name)
  headers = sheet.row(3)
  headers_hash = {}
  headers.each do |h| headers_hash[h] = h if h.is_a? String end

  def actual_cell(hash)
    (hash['Level']!= nil || hash['Top Constraint']!= nil) &&
     hash['Level'] != 'Level'
  end

  arr = []
  sheet.each(headers_hash) do |hash|
    arr << hash if actual_cell(hash)
  end

  arr
end

arr  = sheet_to_arr('Dim - Doors')
arr2 = sheet_to_arr('Dim - MMD')