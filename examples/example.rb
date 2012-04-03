#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

#A Simple Workbook

wb.add_worksheet(:name => "Basic Worksheet") do |sheet|
  sheet.add_row ["First Column", "Second", "Third"]
  sheet.add_row [1, 2, 3]
end

#Using Custom Styles

wb.styles do |s|
  black_cell = s.add_style :bg_color => "00", :fg_color => "FF", :sz => 14, :alignment => { :horizontal=> :center }
  blue_cell =  s.add_style  :bg_color => "0000FF", :fg_color => "FF", :sz => 20, :alignment => { :horizontal=> :center }
  wb.add_worksheet(:name => "Custom Styles") do |sheet|
    sheet.add_row ["Text Autowidth", "Second", "Third"], :style => [black_cell, blue_cell, black_cell]
    sheet.add_row [1, 2, 3], :style => Axlsx::STYLE_THIN_BORDER
  end
end

#Using Custom Border Styles

wb.styles do |s|
  red_border =  s.add_style :border => {:style=>:thin, :color =>"FFFF0000"}
  blue_border =  s.add_style :border => {:style=>:thin, :color =>"FF0000FF"}

  wb.add_worksheet(:name => "Custom Borders") do |sheet|
    sheet.add_row ["wrap", "me", "Up in Red"], :style => red_border
    sheet.add_row [1, 2, 3], :style => blue_border
  end
end


##Using Custom Formatting and date1904
require 'date'
wb.styles do |s|
  date = s.add_style(:format_code => "yyyy-mm-dd", :border => Axlsx::STYLE_THIN_BORDER)
  padded = s.add_style(:format_code => "00#", :border => Axlsx::STYLE_THIN_BORDER)
  percent = s.add_style(:format_code => "0000%", :border => Axlsx::STYLE_THIN_BORDER)
  # wb.date1904 = true # Use the 1904 date system (Used by Excel for Mac < 2011)
  wb.add_worksheet(:name => "Formatting Data") do |sheet|
    sheet.add_row ["Custom Formatted Date", "Percent Formatted Float", "Padded Numbers"], :style => Axlsx::STYLE_THIN_BORDER
    sheet.add_row [Date::strptime('2012-01-19','%Y-%m-%d'), 0.2, 32], :style => [date, percent, padded]
  end
end


##Add an Image

wb.add_worksheet(:name => "Images") do |sheet|
  img = File.expand_path('../image1.jpeg', __FILE__)
  sheet.add_image(:image_src => img, :noSelect => true, :noMove => true) do |image|
    image.width=720
    image.height=666
    image.start_at 2, 2
  end
end

##Add an Image with a hyperlink

wb.add_worksheet(:name => "Image with Hyperlink") do |sheet|
  img = File.expand_path('../image1.jpeg', __FILE__)
  sheet.add_image(:image_src => img, :noSelect => true, :noMove => true, :hyperlink=>"http://axlsx.blogspot.com") do |image|
    image.width=720
    image.height=666
    image.hyperlink.tooltip = "Labeled Link"
    image.start_at 2, 2
  end
end

##Asian Language Support

wb.add_worksheet(:name => "日本語でのシート名") do |sheet|
  sheet.add_row ["日本語"]
  sheet.add_row ["华语/華語"]
  sheet.add_row ["한국어/조선말"]
end

##Styling Columns

wb.styles do |s|
  percent = s.add_style :num_fmt => 9
  wb.add_worksheet(:name => "Styling Columns") do |sheet|
    sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4']
    sheet.add_row [1, 2, 0.3, 4]
    sheet.add_row [1, 2, 0.2, 4]
    sheet.add_row [1, 2, 0.1, 4]
    sheet.col_style 2, percent, :row_offset => 1
  end
end

##Hiding Columns

wb.styles do |s|
  percent = s.add_style :num_fmt => 9
  wb.add_worksheet(:name => "Hidden Column") do |sheet|
    sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4']
    sheet.add_row [1, 2, 0.3, 4]
    sheet.add_row [1, 2, 0.2, 4]
    sheet.add_row [1, 2, 0.1, 4]
    sheet.col_style 2, percent, :row_offset => 1
    sheet.column_info[1].hidden = true
  end
end

##Styling Rows

wb.styles do |s|
  head = s.add_style :bg_color => "00", :fg_color => "FF"
  percent = s.add_style :num_fmt => 9
  wb.add_worksheet(:name => "Styling Rows") do |sheet|
    sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4']
    sheet.add_row [1, 2, 0.3, 4]
    sheet.add_row [1, 2, 0.2, 4]
    sheet.add_row [1, 2, 0.1, 4]
    sheet.col_style 2, percent, :row_offset => 1
    sheet.row_style 0, head
  end
end

##Styling Cell Overrides

wb.add_worksheet(:name => "Cell Level Style Overrides") do |sheet|
  # cell level style overides when adding cells
  sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4'], :sz => 16
  sheet.add_row [1, 2, 3, "=SUM(A2:C2)"]
  # cell level style overrides via sheet range
  sheet["A1:D1"].each { |c| c.color = "FF0000"}
  sheet['A1:D2'].each { |c| c.style = Axlsx::STYLE_THIN_BORDER }
end

##Using formula

wb.add_worksheet(:name => "Using Formulas") do |sheet|
  sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4']
  sheet.add_row [1, 2, 3, "=SUM(A2:C2)"]
end

##Automatic cell types

wb.add_worksheet(:name => "Automatic cell types") do |sheet|
  sheet.add_row ["Date", "Time", "String", "Boolean", "Float", "Integer"]
  sheet.add_row [Date.today, Time.now, "value", true, 0.1, 1]
end

##Merging Cells.

wb.add_worksheet(:name => 'Merging Cells') do |sheet|
  # cell level style overides when adding cells
  sheet.add_row ["col 1", "col 2", "col 3", "col 4"], :sz => 16
  sheet.add_row [1, 2, 3, "=SUM(A2:C2)"]
  sheet.add_row [2, 3, 4, "=SUM(A3:C3)"]
  sheet.add_row ["total", "", "", "=SUM(D2:D3)"]
  sheet.merge_cells("A4:C4")
  sheet["A1:D1"].each { |c| c.color = "FF0000"}
  sheet["A1:D4"].each { |c| c.style = Axlsx::STYLE_THIN_BORDER }
end

##Generating A Bar Chart

wb.add_worksheet(:name => "Bar Chart") do |sheet|
  sheet.add_row ["A Simple Bar Chart"]
  sheet.add_row ["First", "Second", "Third"]
  sheet.add_row [1, 2, 3]
  sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A4", :end_at => "F17") do |chart|
    chart.add_series :data => sheet["A3:C3"], :labels => sheet["A2:C2"], :title => sheet["A1"]
  end
end

##Hide Gridlines in chart
wb.add_worksheet(:name => "Chart With No Gridlines") do |sheet|
  sheet.add_row ["A Simple Bar Chart"]
  sheet.add_row ["First", "Second", "Third"]
  sheet.add_row [1, 2, 3]
  sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A4", :end_at => "F17") do |chart|
    chart.add_series :data => sheet["A3:C3"], :labels => sheet["A2:C2"], :title => sheet["A1"]
    chart.valAxis.gridlines = false
    chart.catAxis.gridlines = false
  end
end

##Generating A Pie Chart

wb.add_worksheet(:name => "Pie Chart") do |sheet|
  sheet.add_row ["First", "Second", "Third", "Fourth"]
  sheet.add_row [1, 2, 3, "=PRODUCT(A2:C2)"]
  sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,2], :end_at => [5, 15], :title => "example 3: Pie Chart") do |chart|
    chart.add_series :data => sheet["A2:D2"], :labels => sheet["A1:D1"]
  end
end

##Data over time

wb.add_worksheet(:name=>'Charting Dates') do |sheet|
   # cell level style overides when adding cells
  sheet.add_row ['Date', 'Value'], :sz => 16
  sheet.add_row [Time.now - (7*60*60*24), 3]
  sheet.add_row [Time.now - (6*60*60*24), 7]
  sheet.add_row [Time.now - (5*60*60*24), 18]
  sheet.add_row [Time.now - (4*60*60*24), 1]
  sheet.add_chart(Axlsx::Bar3DChart) do |chart|
    chart.start_at "B7"
    chart.end_at "H27"
    chart.add_series(:data => sheet["B2:B5"], :labels => sheet["A2:A5"], :title => sheet["B1"])
  end
end

##Generating A Line Chart

wb.add_worksheet(:name => "Line Chart") do |sheet|
  sheet.add_row ["First", 1, 5, 7, 9]
  sheet.add_row ["Second", 5, 2, 14, 9]
  sheet.add_chart(Axlsx::Line3DChart, :title => "example 6: Line Chart", :rotX => 30, :rotY => 20) do |chart|
    chart.start_at 0, 2
    chart.end_at 10, 15
    chart.add_series :data => sheet["B1:E1"], :title => sheet["A1"]
    chart.add_series :data => sheet["B2:E2"], :title => sheet["A2"]
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
    chart.add_series :xData => sheet["B1:E1"], :yData => sheet["B2:E2"], :title => sheet["A1"]
    chart.add_series :xData => sheet["B3:E3"], :yData => sheet["B4:E4"], :title => sheet["A3"]
  end
end

##Auto Filter

wb.add_worksheet(:name => "Auto Filter") do |sheet|
  sheet.add_row ["Build Matrix"]
  sheet.add_row ["Build", "Duration", "Finished", "Rvm"]
  sheet.add_row ["19.1", "1 min 32 sec", "about 10 hours ago", "1.8.7"]
  sheet.add_row ["19.2", "1 min 28 sec", "about 10 hours ago", "1.9.2"]
  sheet.add_row ["19.3", "1 min 35 sec", "about 10 hours ago", "1.9.3"]
  sheet.auto_filter = "A2:D5"
end

##Tables

 wb.add_worksheet(:name => "Table") do |sheet|
   sheet.add_row ["Build Matrix"]
   sheet.add_row ["Build", "Duration", "Finished", "Rvm"]
   sheet.add_row ["19.1", "1 min 32 sec", "about 10 hours ago", "1.8.7"]
   sheet.add_row ["19.2", "1 min 28 sec", "about 10 hours ago", "1.9.2"]
   sheet.add_row ["19.3", "1 min 35 sec", "about 10 hours ago", "1.9.3"]
   sheet.add_table "A2:D5", :name => 'Build Matrix'
 end


##Specifying Column Widths

wb.add_worksheet(:name => "custom column widths") do |sheet|
  sheet.add_row ["I use autowidth and am very wide", "I use a custom width and am narrow"]
  sheet.add_row ['abcdefg', 'This is a very long text and should flow into the right cell', nil, 'xxx' ]
  sheet.column_widths nil, 3, 5, nil
end

##Fit to page printing

wb.add_worksheet(:name => "fit to page") do |sheet|
  sheet.add_row ['this all goes on one page']
  sheet.fit_to_page = true
end


##Hide Gridlines in worksheet
wb.add_worksheet(:name => "No Gridlines") do |sheet|
  sheet.add_row ["This", "Sheet", "Hides", "Gridlines"]
  sheet.show_gridlines = false
end

##Specify Page Margins for printing
margins = {:left => 3, :right => 3, :top => 1.2, :bottom => 1.2, :header => 0.7, :footer => 0.7}
wb.add_worksheet(:name => "print margins", :page_margins => margins) do |sheet|
  sheet.add_row ["this sheet uses customized page margins for printing"]
end

##Validate and Serialize

#p.validate.each { |e| puts e.message }
p.serialize("example.xlsx")

s = p.to_stream()
File.open('example_streamed.xlsx', 'w') { |f| f.write(s.read) }


##Using Shared Strings
p.use_shared_strings = true
p.serialize("shared_strings_example.xlsx")


##Disabling Autowidth
p = Axlsx::Package.new
p.use_autowidth = false
wb = p.workbook
wb.add_worksheet(:name => "No Magick") do | sheet |
  sheet.add_row ['oh look! no autowidth - and no magick loaded in your process']
end
p.validate.each { |e| puts e.message }
p.serialize("no-use_autowidth.xlsx")



