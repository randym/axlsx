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

# secondary axis in line chart
 wb.add_worksheet(:name => "Secondary axis") do |sheet|
  sheet.add_row %w(first second)
  10.times do
    sheet.add_row [ rand(24)+1, rand(24)*100+1]
  end
  sheet.add_chart(Axlsx::LineChart, :title => "Simple Line Chart", :rotX => 30, :rotY => 20) do |chart|
    chart.start_at 0, 5
    chart.end_at 10, 25
    chart.add_series :data => sheet["A2:A11"], :title => sheet["A1"], :color => "5B9BD5", :show_marker => true, :smooth => true
    chart.add_series :data => sheet["B2:B11"], :title => sheet["B1"], :color => "ED7D31", :on_primary_axis => false

    chart.catAxis.title = 'X Axis'
    chart.valAxis.title = 'Primary Axis'
    chart.secValAxis.title = "Secondary Axis"

    # Set the text color of the title
    chart.catAxis.title.text_color = "404040"
    chart.valAxis.title.text_color = "5B9BD5"
    chart.secValAxis.title.text_color = "ED7D31"

    # Set the color of the axis values
    chart.valAxis.text_color = "5B9BD5"
    chart.secValAxis.text_color = "ED7D31"

    # Set the luminance
    chart.catAxis.gridlines_luminance = 0.25
    chart.valAxis.gridlines_luminance = 0.25
  end
 end

p.serialize('basic_charts.xlsx')

