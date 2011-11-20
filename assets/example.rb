require 'rubygems'
require 'axlsx'

p = Axlsx::Package.new do |package|
  package.workbook.add_worksheet do |sheet|
    sheet.add_row ["First", "Second", "Third"], :style => Axlsx::STYLE_THIN_BORDER
    sheet.add_row [1, 2, 3], :style => Axlsx::STYLE_THIN_BORDER
    sheet.add_chart(Axlsx::Bar3DChart, :start_at => [0,2], :end_at => [5, 15], :title=>"example 1: Chart") do |chart|
      chart.add_series :data=>sheet.rows.last.cells, :labels=> sheet.rows.first.cells
    end
  end
  package.serialize("example1.xlsx")
end

p = Axlsx::Package.new
wb = p.workbook

styles = wb.styles

header_style = styles.add_style(:bg_color=>"FF000000", :fg_color=>"FFFFFFFF", :sz=>14, :alignment=>{:horizontal=>:center})
table_title = styles.add_style :sz=>22
date_time = styles.add_style(:num_fmt=>Axlsx::NUM_FMT_YYYYMMDDHHMMSS)
date = styles.add_style :num_fmt=>Axlsx::NUM_FMT_YYYYMMDD
bordered = styles.add_style :border=>Axlsx::STYLE_THIN_BORDER

wb.add_worksheet do |sheet|
  sheet.name = "Example 1"
  # add cells as a row
  sheet.add_row ['Example 1 Styling and Bar Chart'], :style=>table_title
  sheet.add_row ["Formatted Time Stamp", Time.now],:style=>[nil, date_time]
  # adding cells one at a time
  sheet.add_row do |row|
    row.add_cell "Formatted Date" 
    row.add_cell Time.now, :style => date, :type=>:time
  end

  # Bar 3D Chart
  sheet.add_row
  label_row = sheet.add_row ["0 ~ 10", "11 ~ 20", "21 ~ 60"], :style=>header_style
  data_row = sheet.add_row [12, 60, 25], :style=>bordered
  sheet.add_chart(Axlsx::Bar3DChart, :show_legend=>false) do |chart|
    chart.title = Axlsx::Title.new(sheet.rows.first.cells.first)
    chart.add_series :data => data_row.cells, :labels => label_row.cells
    chart.start_at.coord(0, 7)
    chart.end_at.coord(3, 17)
  end

end

# Blocks are not required
percent = styles.add_style :num_fmt=>Axlsx::NUM_FMT_PERCENT, :border=>Axlsx::STYLE_THIN_BORDER
wb.add_worksheet(:name=>"Example 2") do |sheet|
  sheet.add_row ["Example 2 - Styling and Pie Chart"], :style=>table_title
  label_row = sheet.add_row(["Summer", "Fall", "Winter", "Spring"], :style=>header_style)
  data_row =  sheet.add_row [0.5, 0.2, 0.2, 0.1], :style=>percent
  chart = sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,4], :end_at => [4,14], :title => "Pie Consumption per Season") do |chart|
    chart.add_series  :data => data_row.cells, :labels=>label_row.cells
  end
end
# Charts can be build with out data in the sheet
wb.add_worksheet(:name=>"Example 3") do |ws|
  ws.add_row ["Charts can be build without any data in the worksheet"]
  ws.add_chart(Axlsx::Pie3DChart, :title=>"free chart 1") do |chart|
    chart.start_at.coord(0,2)
    chart.end_at.coord(3,12)
    chart.add_series :data => [13,54,1], :labels=>["nothing","in","the sheet"]
  end
  ws.add_chart(Axlsx::Pie3DChart) do |chart|
    chart.start_at.coord(0,13)
    chart.end_at.coord(3,23)
    chart.add_series :data => [1,4,5], :labels=>["nothing","in","the sheet"], :title=>"free chart 2"
  end
end

# wb.add_worksheet(:name=>"Example 4") do |sheet|
#   sheet.add_row ["lots of data!"]  
#   (1..100).each do |i|
#     cells = (1..50).map { |i| rand(5 ** 10) }
#     sheet.add_row cells
#   end
# end

errors = p.validate
if errors.empty?
  p.serialize('test.xlsx')
else
  puts errors.to_s
end

