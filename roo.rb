# https://github.com/roo-rb/roo
require 'roo'

xl = Roo::Spreadsheet.open('src.xlsx')

sheet = xl.sheet(0)
sheet = xl.sheet('Dim - Doors')

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

arr[0]