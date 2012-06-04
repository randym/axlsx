$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
p = Axlsx::Package.new
wb = p.workbook
##Generating A Pie Chart
#```ruby
wb.add_worksheet(:name => "Pie Chart") do |sheet|
  sheet.add_row ["First", "Second", "Third", "Fourth"]
  sheet.add_row [1, 2, 3, '=sum(A2:C2)']
  sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,2], :end_at => [5, 15], :title => "Chart") do |chart|
    chart.add_series :data => sheet["A2:D2"], :labels => sheet["A1:D1"]
  end
end

p.serialize('pie_chart.xlsx')


#
#  Line Chart

p = Axlsx::Package.new
wb = p.workbook
wb.add_worksheet(:name => "Line Chart") do |sheet|
  sheet.add_row ['first', 'second', 'third', 'fourth']
  sheet.add_row [1, 2, 3, '=sum(A2:C2)']
  sheet.add_chart(Axlsx::Line3DChart, :start_at => [0,2], :end_at => [5, 15], :title => "Chart") do |chart|
    chart.add_series :data => sheet["A2:D2"], :labels => sheet["A1:D1"], :title => 'bob'
  end
end
p.serialize('line_chart.xlsx')
##
