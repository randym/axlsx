#!/usr/bin/env ruby
# -*- coding: utf-8 -*-
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

##Using Custom Formatting and date1904
     require 'date'
     wb.styles do |s|
       date = s.add_style(:format_code => "yyyy-mm-dd", :border => Axlsx::STYLE_THIN_BORDER)
       padded = s.add_style(:format_code => "00#", :border => Axlsx::STYLE_THIN_BORDER)
       percent = s.add_style(:format_code => "0000%", :border => Axlsx::STYLE_THIN_BORDER)
       wb.date1904 = true # required for generation on mac
       wb.add_worksheet(:name => "Formatting Data") do |sheet|
         sheet.add_row ["Custom Formatted Date", "Percent Formatted Float", "Padded Numbers"], :style => Axlsx::STYLE_THIN_BORDER
         sheet.add_row [Date::strptime('2012-01-19','%Y-%m-%d'), 0.2, 32], :style => [date, percent, padded]
       end
     end

##Add an Image

     wb.add_worksheet(:name => "Images") do |sheet|
       img = File.expand_path('examples/image1.jpeg') 
       sheet.add_image(:image_src => img, :noSelect => true, :noMove => true) do |image|
         image.width=720
         image.height=666
         image.start_at 2, 2
       end
     end  

##Add an Image with a hyperlink

     wb.add_worksheet(:name => "Image with Hyperlink") do |sheet|
       img = File.expand_path('examples/image1.jpeg') 
       sheet.add_image(:image_src => img, :noSelect => true, :noMove => true, :hyperlink=>"http://axlsx.blogspot.com") do |image|
         image.width=720
         image.height=666
         image.hyperlink.tooltip = "Labeled Link"
         image.start_at 2, 2
       end
     end  

##Asian Language Support

     wb.add_worksheet(:name => "Unicode Support") do |sheet|
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

##Auto Filter

     wb.add_worksheet(:name => "Auto Filter") do |sheet|
       sheet.add_row ["Build Matrix"]
       sheet.add_row ["Build", "Duration", "Finished", "Rvm"]
       sheet.add_row ["19.1", "1 min 32 sec", "about 10 hours ago", "1.8.7"]
       sheet.add_row ["19.2", "1 min 28 sec", "about 10 hours ago", "1.9.2"]
       sheet.add_row ["19.3", "1 min 35 sec", "about 10 hours ago", "1.9.3"]
       sheet.auto_filter = "A2:D5"
     end  

##Validate and Serialize

     p.validate.each { |e| puts e.message }
     p.serialize("example.xlsx")

     

