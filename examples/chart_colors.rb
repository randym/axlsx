#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

##Generating A Bar Chart

wb.add_worksheet(:name => "Bar Chart") do |sheet|
  sheet.add_row ["A Simple Bar Chart"]
  sheet.add_row ["First", "Second", "Third"]
  sheet.add_row [1,2,3.5]
  sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A4", :end_at => "F17", :title => sheet["A1"]) do |chart|
    chart.add_series :data => sheet["A3:C3"], :labels => sheet["A2:C2"], :colors => ['FF0000', '00FF00', '0000FF'], :color => "000000"
    chart.valAxis.label_rotation = -45
    chart.catAxis.label_rotation = 45
  end
end

##Generating A Bar Chart without data in the sheet

wb.add_worksheet(:name => "Hard Bar Chart") do |sheet|
  sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A4", :end_at => "F17", :title => "hard code chart") do |chart|
    chart.add_series :data => [1,3,5], :labels => ['a', 'b', 'c'], :colors => ['FF0000', '00FF00', '0000FF']
    chart.valAxis.label_rotation = -45
    chart.catAxis.label_rotation = 45
  end
end


##Generating A Pie Chart

 wb.add_worksheet(:name => "Pie Chart") do |sheet|
   sheet.add_row ["First", "Second", "Third", "Fourth"]
   sheet.add_row [1, 2, 3, "=PRODUCT(A2:C2)"]
  sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,2], :end_at => [5, 15], :title => "example 3: Pie Chart") do |chart|
     chart.add_series :data => sheet["A2:D2"], :labels => sheet["A1:D1"], :colors => ['FF0000', '00FF00', '0000FF', '000000']
   end
 end

# Line Chart

wb.add_worksheet(:name => "Line Chart") do |sheet|
  sheet.add_row ["First", 1, 5, 7, 9]
  sheet.add_row ["Second", 5, 2, 14, 9]
  sheet.add_chart(Axlsx::Line3DChart, :title => "example 6: Line Chart", :rotX => 30, :rotY => 30) do |chart|
    chart.start_at 0, 2
    chart.end_at 10, 15
    chart.add_series :data => sheet["B1:E1"], :title => sheet["A1"], :color => "FF0000"
    chart.add_series :data => sheet["B2:E2"], :title => sheet["A2"], :color => "00FF00"
  end
end

wb.add_worksheet(:name => 'Line Chart with Axis colors') do |sheet|
  sheet.add_row ["First", 1, 5, 7, 9]
  sheet.add_row ["Second", 5, 2, 14, 9]
  sheet.add_chart(Axlsx::Line3DChart, :title => "example 7: Flat Line Chart", :rot_x => 0, :perspective => 0) do |chart|
    chart.valAxis.color = "FFFF00"
    chart.catAxis.color = "00FFFF"
    chart.serAxis.delete = true
    chart.start_at 0, 2
    chart.end_at 10, 15
    chart.add_series :data => sheet["B1:E1"], :title => sheet["A1"], :color => "FF0000"
    chart.add_series :data => sheet["B2:E2"], :title => sheet["A2"], :color => "00FF00"
  end

end

##Generating A Scatter Chart

 wb.add_worksheet(:name => "Scatter Chart") do |sheet|
   sheet.add_row ["First",  1,  5,  7,  9]
   sheet.add_row ["",       1, 25, 49, 81]
   sheet.add_row ["Second", 5,  2, 14,  9]
   sheet.add_row ["",       5, 10, 15, 20]
   sheet.add_chart(Axlsx::ScatterChart, :title => "example 7: Scatter Chart") do |chart|
     chart.start_at 0, 4
     chart.end_at 10, 19
     chart.add_series :xData => sheet["B1:E1"], :yData => sheet["B2:E2"], :title => sheet["A1"], :color => '00FF00'
     chart.add_series :xData => sheet["B3:E3"], :yData => sheet["B4:E4"], :title => sheet["A3"], :color => 'FF0000'
   end
 end


p.validate.each { |e| puts e.message }
p.serialize 'chart_colors.xlsx'
