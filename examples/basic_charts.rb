$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
p = Axlsx::Package.new
wb = p.workbook

# Pie Chart
wb.add_worksheet(:name => "Pie Chart") do |sheet|
  sheet.add_row ["First", "Second", "Third", "Fourth"]
  sheet.add_row [1, 2, 3, 4]
  sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,2], :end_at => [5, 15], :title=> 'dark corner here') do |chart|
    chart.add_series :data => sheet["A2:D2"], :labels => sheet["A1:D1"]
    chart.d_lbls.show_val = true
    chart.d_lbls.show_percent = true
    chart.d_lbls.d_lbl_pos = :outEnd
    chart.d_lbls.show_leader_lines = true
  end
end

# line chart
 wb.add_worksheet(:name => "Line Chart") do |sheet|
  sheet.add_row ['1', '2', '3', '4']
  sheet.add_row [1, 2, 3, '=sum(A2:C2)']
  sheet.add_chart(Axlsx::Line3DChart, :start_at => [0,2], :end_at => [5, 15], :title => "Chart") do |chart|
    chart.add_series :data => sheet["A2:D2"], :labels => sheet["A1:D1"], :title => 'bob'
    chart.d_lbls.show_val = true
    chart.d_lbls.show_cat_name = true
    chart.catAxis.tick_lbl_pos = :none

  end
 end

# bar chart
 wb.add_worksheet(:name => "Bar Chart") do |sheet|
  sheet.add_row ["A Simple Bar Chart"]
  sheet.add_row ["First", "Second", "Third"]
  sheet.add_row [1, 2, 3]
  sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A4", :end_at => "F17") do |chart|
    chart.add_series :data => sheet["A3:C3"], :labels => sheet["A2:C2"], :title => sheet["A1"]
    chart.valAxis.label_rotation = -45
    chart.catAxis.label_rotation = 45
    chart.d_lbls.d_lbl_pos = :outEnd
    chart.d_lbls.show_val = true

    chart.catAxis.tick_lbl_pos = :none
  end
 end

# specifying colors and title
wb.add_worksheet(:name => "Colored Pie Chart") do |sheet|
  sheet.add_row ["First", "Second", "Third", "Fourth"]
  sheet.add_row [1, 2, 3, "=PRODUCT(A2:C2)"]
  sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,2], :end_at => [5, 15], :title => "example 3: Pie Chart") do |chart|
    chart.add_series :data => sheet["A2:D2"], :labels => ["A1:D1"], :colors => ['FF0000', '00FF00', '0000FF']
  end
end

p.serialize('basic_charts.xlsx')

