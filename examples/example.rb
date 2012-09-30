#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
# $LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

#```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook
#```

#A Simple Workbook

#```ruby
wb.add_worksheet(:name => "Basic Worksheet") do |sheet|
  sheet.add_row ["First Column", "Second", "Third"]
  sheet.add_row [1, 2, 3]
end
#```

#Using Custom Styles

#```ruby
# Each cell allows a single, predified style.
# When using add_row, the value in the :style array at the same index as the cell's column will be applied to that cell.
# Alternatively, you can apply a style to an entire row by using an integer value for :style.

wb.styles do |s|
  black_cell = s.add_style :bg_color => "00", :fg_color => "FF", :sz => 14, :alignment => { :horizontal=> :center }
  blue_cell =  s.add_style  :bg_color => "0000FF", :fg_color => "FF", :sz => 20, :alignment => { :horizontal=> :center }
  wb.add_worksheet(:name => "Custom Styles") do |sheet|

    # Applies the black_cell style to the first and third cell, and the blue_cell style to the second.
    sheet.add_row ["Text Autowidth", "Second", "Third"], :style => [black_cell, blue_cell, black_cell]

    # Applies the thin border to all three cells
    sheet.add_row [1, 2, 3], :style => Axlsx::STYLE_THIN_BORDER
  end
end
#```

##Styling Cell Overrides

#```ruby
#Some of the style attributes can also be set at the cell level. Cell level styles take precedence over Custom Styles shown in the previous example.

wb.add_worksheet(:name => "Cell Level Style Overrides") do |sheet|

  # this will set the font size for each cell.
  sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4'], :sz => 16

  sheet.add_row [1, 2, 3, "=SUM(A2:C2)"]

  # You can also apply cell style overrides to a range of cells
  sheet["A1:D1"].each { |c| c.color = "FF0000" }
  sheet['A1:D2'].each { |c| c.style = Axlsx::STYLE_THIN_BORDER }
end
##```

##Using Custom Border Styles

#```ruby
#Axlsx defines a thin border style, but you can easily create and use your own.
wb.styles do |s|
  red_border =  s.add_style :border => { :style => :thick, :color =>"FFFF0000", :edges => [:left, :right] }
  blue_border =  s.add_style :border => { :style => :thick, :color =>"FF0000FF"}

  wb.add_worksheet(:name => "Custom Borders") do |sheet|
    sheet.add_row ["wrap", "me", "Up in Red"], :style => red_border
    sheet.add_row [1, 2, 3], :style => blue_border
  end
end
##```


##Styling Rows and Columns

#```ruby
wb.styles do |s|
  head = s.add_style :bg_color => "00", :fg_color => "FF"
  percent = s.add_style :num_fmt => 9
  wb.add_worksheet(:name => "Columns and Rows") do |sheet|
    sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4', 'col5']
    sheet.add_row [1, 2, 0.3, 4, 5.0]
    sheet.add_row [1, 2, 0.2, 4, 5.0]
    sheet.add_row [1, 2, 0.1, 4, 5.0]

    #apply the percent style to the column at index 2 skipping the first row.
    sheet.col_style 2, percent, :row_offset => 1

    # apply the head style to the first row.
    sheet.row_style 0, head

    #Hide the 5th column
    sheet.column_info[4].hidden = true

    #Set the second column outline level
    sheet.column_info[1].outlineLevel = 2

    sheet.rows[3].hidden = true
    sheet.rows[1].outlineLevel = 2
  end
end
##```


##Specifying Column Widths

#```ruby
wb.add_worksheet(:name => "custom column widths") do |sheet|
  sheet.add_row ["I use autowidth and am very wide", "I use a custom width and am narrow"]
  sheet.add_row ['abcdefg', 'This is a very long text and should flow into the right cell', nil, 'xxx' ]
  sheet.column_widths nil, 3, 5, nil
end
##```

##Merging Cells.

#```ruby
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
##```

##Add an Image with a hyperlink

#```ruby
wb.add_worksheet(:name => "Image with Hyperlink") do |sheet|
  img = File.expand_path('../image1.jpeg', __FILE__)
  # specifying the :hyperlink option will add a hyper link to your image.
  # @note - Numbers does not support this part of the specification.
  sheet.add_image(:image_src => img, :noSelect => true, :noMove => true, :hyperlink=>"http://axlsx.blogspot.com") do |image|
    image.width=720
    image.height=666
    image.hyperlink.tooltip = "Labeled Link"
    image.start_at 2, 2
  end
end
#```

##Using Custom Formatting and date1904

#```ruby
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
#```

##Asian Language Support

#```ruby
wb.add_worksheet(:name => "日本語でのシート名") do |sheet|
  sheet.add_row ["日本語"]
  sheet.add_row ["华语/華語"]
  sheet.add_row ["한국어/조선말"]
end
##```

##Using formula

#```ruby
wb.add_worksheet(:name => "Using Formulas") do |sheet|
  sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4']
  sheet.add_row [1, 2, 3, "=SUM(A2:C2)"]
end
##```

##Auto Filter

#```ruby
wb.add_worksheet(:name => "Auto Filter") do |sheet|
  sheet.add_row ["Build Matrix"]
  sheet.add_row ["Build", "Duration", "Finished", "Rvm"]
  sheet.add_row ["19.1", "1 min 32 sec", "about 10 hours ago", "1.8.7"]
  sheet.add_row ["19.2", "1 min 28 sec", "about 10 hours ago", "1.9.2"]
  sheet.add_row ["19.3", "1 min 35 sec", "about 10 hours ago", "1.9.3"]
  sheet.auto_filter = "A2:D5"
  sheet.auto_filter.add_column 3, :filters, :filter_items => ['1.9.2', '1.8.7']
end
#```

##Automatic cell types


#```ruby
wb.add_worksheet(:name => "Automatic cell types") do |sheet|
  date_format = wb.styles.add_style :format_code => 'YYYY-MM-DD'
  time_format = wb.styles.add_style :format_code => 'hh:mm:ss'
  sheet.add_row ["Date", "Time", "String", "Boolean", "Float", "Integer"]
  sheet.add_row [Date.today, Time.now, "value", true, 0.1, 1], :style => [date_format, time_format]
end


# Hyperlinks in worksheet
wb.add_worksheet(:name => 'hyperlinks') do |sheet|
  # external references
  sheet.add_row ['axlsx']
  sheet.add_hyperlink :location => 'https://github.com/randym/axlsx', :ref => sheet.rows.first.cells.first
  # internal references
  sheet.add_hyperlink :location => "'Next Sheet'!A1", :ref => 'A2', :target => :sheet
  sheet.add_row ['next sheet']
end

wb.add_worksheet(:name => 'Next Sheet') do |sheet|
  sheet.add_row ['hello!']
end
###```

##Number formatting and currency
wb.add_worksheet(:name => "Formats and Currency") do |sheet|
  currency = wb.styles.add_style :num_fmt => 5
  red_negative = wb.styles.add_style :num_fmt => 8
  comma = wb.styles.add_style :num_fmt => 3
  super_funk = wb.styles.add_style :format_code => '[Green]#'
  sheet.add_row %w(Currency RedNegative Comma Custom)
  sheet.add_row [1500, -122.34, 123456789, 594829], :style=> [currency, red_negative, comma, super_funk]
end

##Generating A Bar Chart

#```ruby
wb.add_worksheet(:name => "Bar Chart") do |sheet|
  sheet.add_row ["A Simple Bar Chart"]
  %w(first second third).each { |label| sheet.add_row [label, rand(24)+1] }
  sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A6", :end_at => "F20") do |chart|
    chart.add_series :data => sheet["B2:B4"], :labels => sheet["A2:A4"], :title => sheet["A1"]
  end
end
##```

##Hide Gridlines in chart

#```ruby
wb.add_worksheet(:name => "Chart With No Gridlines") do |sheet|
  sheet.add_row ["Bar Chart without gridlines"]
  %w(first second third).each { |label| sheet.add_row [label, rand(24)+1] }
  sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A6", :end_at => "F20") do |chart|
    chart.add_series :data => sheet["B2:B4"], :labels => sheet["A2:A4"]
    chart.valAxis.gridlines = false
    chart.catAxis.gridlines = false
  end
end
#```

##Generating A Pie Chart

#```ruby
wb.add_worksheet(:name => "Pie Chart") do |sheet|
  sheet.add_row ["Simple Pie Chart"]
  %w(first second third).each { |label| sheet.add_row [label, rand(24)+1] }
  sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,5], :end_at => [10, 20], :title => "example 3: Pie Chart") do |chart|
    chart.add_series :data => sheet["B2:B4"], :labels => sheet["A2:A4"],  :colors => ['FF0000', '00FF00', '0000FF']
  end
end
#```

##Generating A Line Chart

#```ruby
wb.add_worksheet(:name => "Line Chart") do |sheet|
  sheet.add_row ["Simple Line Chart"]
  sheet.add_row %w(first second)
  4.times do
    sheet.add_row [ rand(24)+1, rand(24)+1]
  end
  sheet.add_chart(Axlsx::Line3DChart, :title => "Simple Line Chart", :rotX => 30, :rotY => 20) do |chart|
    chart.start_at 0, 5
    chart.end_at 10, 20
    chart.add_series :data => sheet["A3:A6"], :title => sheet["A2"]
    chart.add_series :data => sheet["B3:B6"], :title => sheet["B2"]
    chart.catAxis.title = 'X Axis'
    chart.valAxis.title = 'Y Axis'
  end
end
#```

##Generating A Scatter Chart

#```ruby
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
#```


##Tables

#```ruby
wb.add_worksheet(:name => "Table") do |sheet|
  sheet.add_row ["Build Matrix"]
  sheet.add_row ["Build", "Duration", "Finished", "Rvm"]
  sheet.add_row ["19.1", "1 min 32 sec", "about 10 hours ago", "1.8.7"]
  sheet.add_row ["19.2", "1 min 28 sec", "about 10 hours ago", "1.9.2"]
  sheet.add_row ["19.3", "1 min 35 sec", "about 10 hours ago", "1.9.3"]
  sheet.add_table "A2:D5", :name => 'Build Matrix', :style_info => { :name => "TableStyleMedium23" }
end
#```


##Fit to page printing

#```ruby
wb.add_worksheet(:name => "fit to page") do |sheet|
  sheet.add_row ['this all goes on one page']
  sheet.fit_to_page = true
end
##```


##Hide Gridlines in worksheet

#```ruby
wb.add_worksheet(:name => "No Gridlines") do |sheet|
  sheet.add_row ["This", "Sheet", "Hides", "Gridlines"]
  sheet.show_gridlines = false
end
##```

#```ruby
wb.add_worksheet(:name => "repeated header") do |sheet|
  sheet.add_row %w(These Column Header Will Render On Every Printed Sheet)
  200.times { sheet.add_row %w(1 2 3 4 5 6 7 8) }
  wb.add_defined_name("'repeated header'!$1:$1", :local_sheet_id => sheet.index, :name => '_xlnm.Print_Titles') 
end

# Sheet Protection and excluding cells from locking.
unlocked = wb.styles.add_style :locked => false
wb.add_worksheet(:name => 'Sheet Protection') do |sheet|
  sheet.sheet_protection.password = 'fish'
  sheet.add_row [1, 2 ,3] # These cells will be locked
  sheet.add_row [4, 5, 6], :style => unlocked # these cells will not!
end


##Specify page margins and other options for printing

#```ruby
margins = {:left => 3, :right => 3, :top => 1.2, :bottom => 1.2, :header => 0.7, :footer => 0.7}
setup = {:fit_to_width => 1, :orientation => :landscape, :paper_width => "297mm", :paper_height => "210mm"}
options = {:grid_lines => true, :headings => true, :horizontal_centered => true}
wb.add_worksheet(:name => "print margins", :page_margins => margins, :page_setup => setup, :print_options => options) do |sheet|
  sheet.add_row ["this sheet uses customized print settings"]
end
#```

## Add Comments to your spreadsheet 
#``` ruby
wb.add_worksheet(:name => 'comments') do |sheet|
  sheet.add_row ['Can we build it?']
  sheet.add_comment :ref => 'A1', :author => 'Bob', :text => 'Yes We Can!'
end

## Frozen/Split panes
## ``` ruby
wb.add_worksheet(:name => 'fixed headers') do |sheet|
 sheet.add_row(['',  (0..99).map { |i| "column header #{i}" }].flatten )
  100.times.with_index { |index| sheet << ["row header", (0..index).to_a].flatten }
  sheet.sheet_view.pane do |pane|
    pane.top_left_cell = "B2"
    pane.state = :frozen_split
    pane.y_split = 1
    pane.x_split = 1
    pane.active_pane = :bottom_right
  end
end

# conditional formatting
#
percent = wb.styles.add_style(:format_code => "0.00%", :border => Axlsx::STYLE_THIN_BORDER)
money = wb.styles.add_style(:format_code => '0,000', :border => Axlsx::STYLE_THIN_BORDER)

# define the style for conditional formatting
profitable = wb.styles.add_style( :fg_color=>"FF428751",
                                    :type => :dxf)

wb.add_worksheet(:name => "Conditional Cell Is") do |ws|

  # Generate 20 rows of data
  ws.add_row ["Previous Year Quarterly Profits (JPY)"]
  ws.add_row ["Quarter", "Profit", "% of Total"]
  offset = 3
  rows = 20
  offset.upto(rows + offset) do |i|
    ws.add_row ["Q#{i}", 10000*((rows/2-i) * (rows/2-i)), "=100*B#{i}/SUM(B3:B#{rows+offset})"], :style=>[nil, money, percent]
  end

# Apply conditional formatting to range B3:B100 in the worksheet
  ws.add_conditional_formatting("B3:B100", { :type => :cellIs, :operator => :greaterThan, :formula => "100000", :dxfId => profitable, :priority => 1 })
end

wb.add_worksheet(:name => "Conditional Color Scale") do |ws|
  ws.add_row ["Previous Year Quarterly Profits (JPY)"]
  ws.add_row ["Quarter", "Profit", "% of Total"]
  offset = 3
  rows = 20
  offset.upto(rows + offset) do |i|
    ws.add_row ["Q#{i}", 10000*((rows/2-i) * (rows/2-i)), "=100*B#{i}/SUM(B3:B#{rows+offset})"], :style=>[nil, money, percent]
  end
# Apply conditional formatting to range B3:B100 in the worksheet
  color_scale = Axlsx::ColorScale.new
  ws.add_conditional_formatting("B3:B100", { :type => :colorScale, :operator => :greaterThan, :formula => "100000", :dxfId => profitable, :priority => 1, :color_scale => color_scale })
end


wb.add_worksheet(:name => "Conditional Data Bar") do |ws|
  ws.add_row ["Previous Year Quarterly Profits (JPY)"]
  ws.add_row ["Quarter", "Profit", "% of Total"]
  offset = 3
  rows = 20
  offset.upto(rows + offset) do |i|
    ws.add_row ["Q#{i}", 10000*((rows/2-i) * (rows/2-i)), "=100*B#{i}/SUM(B3:B#{rows+offset})"], :style=>[nil, money, percent]
  end
# Apply conditional formatting to range B3:B100 in the worksheet
  data_bar = Axlsx::DataBar.new
  ws.add_conditional_formatting("B3:B100", { :type => :dataBar, :dxfId => profitable, :priority => 1, :data_bar => data_bar })
end

wb.add_worksheet(:name => "Conditional Format Icon Set") do |ws|
  ws.add_row ["Previous Year Quarterly Profits (JPY)"]
  ws.add_row ["Quarter", "Profit", "% of Total"]
  offset = 3
  rows = 20
  offset.upto(rows + offset) do |i|
    ws.add_row ["Q#{i}", 10000*((rows/2-i) * (rows/2-i)), "=100*B#{i}/SUM(B3:B#{rows+offset})"], :style=>[nil, money, percent]
  end
# Apply conditional formatting to range B3:B100 in the worksheet
  icon_set = Axlsx::IconSet.new
  ws.add_conditional_formatting("B3:B100", { :type => :iconSet, :dxfId => profitable, :priority => 1, :icon_set => icon_set })
end

##Validate and Serialize

#```ruby
# Serialize directly to file
p.serialize("example.xlsx")

# or

#Serialize to a stream
s = p.to_stream()
File.open('example_streamed.xlsx', 'w') { |f| f.write(s.read) }
#```

##Using Shared Strings

#```ruby
# This is required by Numbers
p.use_shared_strings = true
p.serialize("shared_strings_example.xlsx")
#```

#p.validate do |er|
  #puts er.inspect
#end
##Disabling Autowidth

#```ruby
p = Axlsx::Package.new
p.use_autowidth = false
wb = p.workbook
wb.add_worksheet(:name => "Manual Widths") do | sheet |
  sheet.add_row ['oh look! no autowidth']
end
p.validate.each { |e| puts e.message }
p.serialize("no-use_autowidth.xlsx")
#```




