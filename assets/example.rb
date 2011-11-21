     require 'rubygems'
     require 'axlsx'

#A Simple Workbook
      p = Axlsx::Package.new
      p.workbook.add_worksheet do |sheet|
        sheet.add_row ["First", "Second", "Third"]
        sheet.add_row [1, 2, 3]
      end
      p.serialize("example1.xlsx")

#Generating A Bar Chart
     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["First", "Second", "Third"]
       sheet.add_row [1, 2, 3]
       sheet.add_chart(Axlsx::Bar3DChart, :start_at => [0,2], :end_at => [5, 15], :title=>"example 2: Chart") do |chart|
         chart.add_series :data=>sheet.rows.last.cells, :labels=> sheet.rows.first.cells
       end
     end  
     p.serialize("example2.xlsx")

#Generating A Pie Chart
     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["First", "Second", "Third"]
       sheet.add_row [1, 2, 3]
       sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,2], :end_at => [5, 15], :title=>"example 3: Pie Chart") do |chart|
         chart.add_series :data=>sheet.rows.last.cells, :labels=> sheet.rows.first.cells
       end
     end  
     p.serialize("example3.xlsx")


#Using Custom Styles

     p = Axlsx::Package.new
     wb = p.workbook
     black_cell = wb.styles.add_style :bg_color => "FF000000", :fg_color => "FFFFFFFF", :sz=>14, :alignment => { :horizontal=> :center }
     blue_cell = wb.styles.add_style  :bg_color => "FF0000FF", :fg_color => "FFFFFFFF", :sz=>14, :alignment => { :horizontal=> :center }
     wb.add_worksheet do |sheet|
       sheet.add_row ["Text Autowidth", "Second", "Third"], :style => [black_cell, blue_cell, black_cell]
       sheet.add_row [1, 2, 3], :style => Axlsx::STYLE_THIN_BORDER
     end
     p.serialize("example4.xlsx")

#Using Custom Formatting and date1904

     p = Axlsx::Package.new
     wb = p.workbook
     date = wb.styles.add_style :format_code=>"yyyy-mm-dd", :border => Axlsx::STYLE_THIN_BORDER
     padded = wb.styles.add_style :format_code=>"00#", :border => Axlsx::STYLE_THIN_BORDER
     percent = wb.styles.add_style :format_code=>"0%", :border => Axlsx::STYLE_THIN_BORDER
     wb.date1904 = true # required for generation on mac
     wb.add_worksheet do |sheet|
       sheet.add_row ["Custom Formatted Date", "Percent Formatted Float", "Padded Numbers"], :style => Axlsx::STYLE_THIN_BORDER
       sheet.add_row [Time.now, 0.2, 32], :style => [date, percent, padded]
     end
     p.serialize("example5.xlsx")

#Validation
     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["First", "Second", "Third"]
       sheet.add_row [1, 2, 3]
     end

     p.validate.each do |error|
       puts error.inspect
     end
      
#Generating A Line Chart

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["First", 1, 5, 7, 9]
       sheet.add_row ["Second", 5, 2, 14, 9]
       sheet.add_chart(Axlsx::Line3DChart, :start_at => [0,2], :end_at => [10, 15], :title=>"example 6: Line Chart") do |chart|
         chart.add_series :data=>sheet.rows.first.cells[(1..-1)], :title=> sheet.rows.first.cells.first
         chart.add_series :data=>sheet.rows.last.cells[(1..-1)], :title=> sheet.rows.last.cells.first
       end
     end  
     p.serialize("example6.xlsx")
