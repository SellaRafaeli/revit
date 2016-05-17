# 1. pre-requisites: install ruby, gem install roo, gem install spreadsheet
# 2. put this file in a folder with a file called "src.xlsx"
# 3. run 'ruby all_revit_data.rb'
# 4. see target.xlsx

# https://github.com/roo-rb/roo
require 'roo' # reading
require 'spreadsheet' 

# get sheet as hash of items 
def sheet_to_arr(sheet_name)
  xl = Roo::Spreadsheet.open('src.xlsx')
  sheet = xl.sheet(sheet_name)
  headers = sheet.row(3) #headers are on row 3
  headers_hash = {}
  headers.each do |h| headers_hash[h] = h if h.is_a? String end

  def actual_cell(hash)
    (hash['Level']!= nil || hash['Top Constraint']!= nil) &&
     hash['Level'] != 'Level' && 
     hash['Top Constraint'] != 'Top Constraint' 
  end

  arr = []
  sheet.each(headers_hash) do |hash|
    hash['sheet_name'] = sheet_name
    arr << hash if actual_cell(hash)
  end
  arr
end

# concat all items to array of hashes with identical fields
def concat_items(arrs)
  all_keys = arrs.map { |arr| arr[0].keys }.flatten.uniq#.sort
  all_items = arrs.flatten
  all_items.map! {|i| 
    item = {}
    all_keys.each {|key| item[key] = i[key] || nil } 
    item
  }
  return all_items, all_keys
end

# write array of hashes to spreadsheet
def write_to_file(rows = [], headers = [], file_name) #write_to_sheet
  book = Spreadsheet::Workbook.new
  sheet = book.create_worksheet :name => file_name
  sheet.row(0).concat(headers) 
  rows.each_with_index { |row, index| sheet.row(index+1).concat(row.values) } 

  format = Spreadsheet::Format.new(:weight => :bold)
  sheet.row(0).default_format = format
  
  book.write file_name+'.xls'
end

# all data
$doors   = sheet_to_arr('Dim - Doors')
$windows = sheet_to_arr('Dim - Windows')
$multi   = sheet_to_arr('Dim - Multi')
$rooms   = sheet_to_arr('Room Schedule')

$sheets = ['Dim - MMD', 'Dim - Wall Schedule']
$arrs = [$doors] + [$windows] + [$multi] + $sheets.map {|sheet| sheet_to_arr(sheet)}
$items, $keys = concat_items($arrs)
write_to_file($items, $keys, 'All_Data')

# rooms 
def room_items(items_group, room_id)
  items_group.select {|item| item['To Room: Number'] == room_id || item['From Room: Number'] == room_id || item['Room: Number'] == room_id }
end

def room_generic_data(room)  
  name = "Room "+room['Number'].to_s+ ' -- '+room['Name']
  perimeter = 'Perimeter - '+room['Perimeter'].to_s
  area      = 'Area - '+room['Area'].to_s
  general_data    = [{val: name}, {val: perimeter}, {val: area}]
end

def room_data(room_id)
  room = $rooms.select { |r| r['Number'] == room_id }[0]
  return nil unless room 

  empty_line = [{'empty_line': ''}]
  windows = [{val: 'Windows'}] + room_items($windows, room_id)
  doors   = [{val: 'Doors'}]   
  doors_to_room = [{val: 'To Room'}] + $doors.select{|d| d['To Room: Number'] == room_id}
  doors_from_room = [{val: 'From Room'}] + $doors.select{|d| d['From Room: Number'] == room_id}
  items   = [{val: 'Multi Items'}] + room_items($multi, room_id)
  
  res = empty_line + room_generic_data(room) + empty_line + windows + doors + doors_to_room + doors_from_room + items + empty_line
end

$room_lines = (1..20).to_a.map {|room_id| room_data(room_id) }.compact.flatten
write_to_file($room_lines, $keys, 'Room_Data')

##
puts "done"
#require './all_revit_data'