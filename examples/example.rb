#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

#```ruby
require 'axlsx'
examples = []
examples << :basic
examples << :custom_styles
examples << :wrap_text
examples << :cell_style_override
examples << :custom_borders
examples << :surrounding_border
examples << :deep_custom_borders
examples << :row_column_style
examples << :fixed_column_width
examples << :outline_level
examples << :merge_cells
examples << :images
examples << :format_dates
examples << :mbcs
examples << :formula
examples << :auto_filter
examples << :data_types
examples << :override_data_types
examples << :hyperlinks
examples << :number_currency_format
examples << :venezuela_currency
examples << :bar_chart
examples << :chart_gridlines
examples << :pie_chart
examples << :line_chart
examples << :scatter_chart
examples << :tables
examples << :fit_to_page
examples << :hide_gridlines
examples << :repeated_header
examples << :defined_name
examples << :printing
examples << :header_footer
examples << :comments
examples << :panes
examples << :sheet_view
examples << :conditional_formatting
examples << :streaming
examples << :shared_strings
examples << :no_autowidth
examples << :cached_formula
p = Axlsx::Package.new
wb = p.workbook
#```

## A Simple Workbook

#```ruby
if examples.include? :basic
  wb.add_worksheet(:name => "Basic Worksheet") do |sheet|
    sheet.add_row ["First Column", "Second", "Third"]
    sheet.add_row [1, 2, 3]
  end
end
#```

#Using Custom Styles

#```ruby
# Each cell allows a single, predified style.
# When using add_row, the value in the :style array at the same index as the cell's column will be applied to that cell.
# Alternatively, you can apply a style to an entire row by using an integer value for :style.
if examples.include? :custom_styles
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
end


#```ruby
# A simple example of wrapping text. Seems this may not be working in Libre Office so here is an example for me to play with.
if examples.include? :wrap_text
  wb.styles do |s|
    wrap_text = s.add_style :fg_color=> "FFFFFF",
                            :b => true,
                            :bg_color => "004586",
                            :sz => 12,
                            :border => { :style => :thin, :color => "00" },
                            :alignment => { :horizontal => :center,
                                            :vertical => :center ,
                                            :wrap_text => true}
    wb.add_worksheet(:name => 'wrap text') do |sheet|
      sheet.add_row ['Torp, White and Cronin'], :style=>wrap_text
      sheet.column_info.first.width = 5
    end
  end
end

##Styling Cell Overrides

#```ruby
#Some of the style attributes can also be set at the cell level. Cell level styles take precedence over Custom Styles shown in the previous example.
if examples.include? :cell_style_override
  wb.add_worksheet(:name => "Cell Level Style Overrides") do |sheet|

    # this will set the font size for each cell.
    sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4'], :sz => 16

    sheet.add_row [1, 2, 3, "=SUM(A2:C2)"]
    sheet.add_row %w(u shadow sz b i strike outline)
    sheet.rows.last.cells[0].u = :double
    sheet.rows.last.cells[1].shadow = true
    sheet.rows.last.cells[2].sz = 20
    sheet.rows.last.cells[3].b = true
    sheet.rows.last.cells[4].i = true
    sheet.rows.last.cells[5].strike = true
    sheet.rows.last.cells[6].outline = 1
    # You can also apply cell style overrides to a range of cells
    sheet["A1:D1"].each { |c| c.color = "FF0000" }
    sheet['A1:D2'].each { |c| c.style = Axlsx::STYLE_THIN_BORDER }
  end
end
##```

##Using Custom Border Styles

#```ruby
#Axlsx defines a thin border style, but you can easily create and use your own.
if examples.include? :custom_borders
  wb.styles do |s|
    red_border =  s.add_style :border => { :style => :thick, :color =>"FFFF0000", :edges => [:left, :right] }
    blue_border =  s.add_style :border => { :style => :thick, :color =>"FF0000FF"}

    wb.add_worksheet(:name => "Custom Borders") do |sheet|
      sheet.add_row ["wrap", "me", "Up in Red"], :style => red_border
      sheet.add_row [1, 2, 3], :style => blue_border
    end
  end
end

#```ruby
# More Custom Borders
if examples.include? :surrounding_border

  # Stuff like this is why I LOVE RUBY
  # If you dont know about hash default values
  # LEARN IT! LIVE IT! LOVE IT!
  defaults =  { :style => :thick, :color => "000000" }
  borders = Hash.new do |hash, key|
    hash[key] = wb.styles.add_style :border => defaults.merge( { :edges => key.to_s.split('_').map(&:to_sym) } )
  end
  top_row =  [0, borders[:top_left], borders[:top], borders[:top], borders[:top_right]]
  middle_row = [0, borders[:left], nil, nil, borders[:right]]
  bottom_row = [0, borders[:bottom_left], borders[:bottom], borders[:bottom], borders[:bottom_right]]

  wb.add_worksheet(:name => "Surrounding Border") do |ws|
    ws.add_row []
    ws.add_row ['', 1,2,3,4], :style => top_row
    ws.add_row ['', 5,6,7,8], :style => middle_row
    ws.add_row ['', 9, 10, 11, 12]

    #This works too!
    ws.rows.last.style = bottom_row

  end
end

#```ruby
# Hacking border styles
if examples.include? :deep_custom_borders
  wb.styles do |s|
    top_bottom =  s.add_style :border => { :style => :thick, :color =>"FFFF0000", :edges => [:top, :bottom] }
    border = s.borders[s.cellXfs[top_bottom].borderId]
    # edit existing border parts
    border.prs.each do |part|
      case part.name
      when :top
        part.color = Axlsx::Color.new(:rgb => "FFFF0000")
      when :bottom
        part.color = Axlsx::Color.new(:rgb => "FF00FF00")
      end
    end

    border.prs << Axlsx::BorderPr.new(:name => :left, :color => Axlsx::Color.new(:rgb => '0000FF'), :style => :mediumDashed)
    wb.add_worksheet(:name => 'hacked borders') do |sheet|
      sheet.add_row [1,2,3], :style=>top_bottom
    end
  end
end
##```


##Styling Rows and Columns

#```ruby
if examples.include? :row_column_style
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
end
##```


##Specifying Column Widths

#```ruby
if examples.include? :fixed_column_width
  wb.add_worksheet(:name => "custom column widths") do |sheet|
    sheet.add_row ["I use autowidth and am very wide", "I use a custom width and am narrow"]
    sheet.add_row ['abcdefg', 'This is a very long text and should flow into the right cell', nil, 'xxx' ]
    sheet.column_widths nil, 3, 5, nil
  end
end

#```ruby
if examples.include? :outline_level
  wb.add_worksheet(name: 'outline_level') do |sheet|
    sheet.add_row [1, 2, 3, Time.now, 5, 149455.15]
    sheet.add_row [1, 2, 5, 6, 5,14100.19]
    sheet.add_row [9500002267, 1212, 1212, 5,14100.19]
    sheet.outline_level_rows 0, 2
    sheet.outline_level_columns 0, 2
  end
end
##```

##Merging Cells.

#```ruby
if examples.include? :merge_cells
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
end
##```

##Add an Image with a hyperlink

#```ruby
if examples.include? :images
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
end
#```

##Using Custom Formatting and date1904

#```ruby
if examples.include? :format_dates
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
end
#```

##Asian Language Support

#```ruby
if examples.include? :mbcs
  wb.add_worksheet(:name => "日本語でのシート名") do |sheet|
    sheet.add_row ["日本語"]
    sheet.add_row ["华语/華語"]
    sheet.add_row ["한국어/조선말"]
  end
end
##```

##Using formula

#```ruby
if examples.include? :formula
  wb.add_worksheet(:name => "Using Formulas") do |sheet|
    sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4']
    sheet.add_row [1, 2, 3, "=SUM(A2:C2)"]
  end
end
##```

##Auto Filter

#```ruby
if examples.include? :auto_filter
  wb.add_worksheet(:name => "Auto Filter") do |sheet|
    sheet.add_row ["Build Matrix"]
    sheet.add_row ["Build", "Duration", "Finished", "Rvm"]
    sheet.add_row ["19.1", "1 min 32 sec", "about 10 hours ago", "1.8.7"]
    sheet.add_row ["19.2", "1 min 28 sec", "about 10 hours ago", "1.9.2"]
    sheet.add_row ["19.3", "1 min 35 sec", "about 10 hours ago", "1.9.3"]
    sheet.auto_filter = "A2:D5"
    sheet.auto_filter.add_column 3, :filters, :filter_items => ['1.9.2', '1.8.7']
  end
end
#```

##Automatic cell types


#```ruby
if examples.include? :data_types
  wb.add_worksheet(:name => "Automatic cell types") do |sheet|
    date_format = wb.styles.add_style :format_code => 'YYYY-MM-DD'
    time_format = wb.styles.add_style :format_code => 'hh:mm:ss'
    sheet.add_row ["Date", "Time", "String", "Boolean", "Float", "Integer"]
    sheet.add_row [Date.today, Time.now, "value", true, 0.1, 1], :style => [date_format, time_format]
  end
end

#```ruby
if examples.include? :override_data_types
  wb.add_worksheet(:name => "Override Data Type") do |sheet|
    sheet.add_row ['dont eat my zeros!', '0088'] , :types => [nil, :string]
  end
end
# Hyperlinks in worksheet
if examples.include? :hyperlinks
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
end
###```

##Number formatting and currency
if examples.include? :number_currency_format
  wb.add_worksheet(:name => "Formats and Currency") do |sheet|
    currency = wb.styles.add_style :num_fmt => 5
    red_negative = wb.styles.add_style :num_fmt => 8
    comma = wb.styles.add_style :num_fmt => 3
    super_funk = wb.styles.add_style :format_code => '[Green]#'
    sheet.add_row %w(Currency RedNegative Comma Custom)
    sheet.add_row [1500, -122.34, 123456789, 594829], :style=> [currency, red_negative, comma, super_funk]
  end
end

## Venezuala currency
if examples.include? :venezuela_currency
  wb.add_worksheet(:name => 'Venezuala_currency') do |sheet|
    number = wb.styles.add_style :format_code => '#.##0\,00'
    sheet.add_row [2.5] , :style => [number]
  end
end

##Generating A Bar Chart

#```ruby
if examples.include? :bar_chart
  wb.add_worksheet(:name => "Bar Chart") do |sheet|
    sheet.add_row ["A Simple Bar Chart"]
    %w(first second third).each { |label| sheet.add_row [label, rand(24)+1] }
    sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A6", :end_at => "F20") do |chart|
      chart.add_series :data => sheet["B2:B4"], :labels => sheet["A2:A4"], :title => sheet["A1"], :colors => ["00FF00", "0000FF"]
    end
  end
end

##```

##Hide Gridlines in chart

#```ruby
if examples.include? :chart_gridlines
  wb.add_worksheet(:name => "Chart With No Gridlines") do |sheet|
    sheet.add_row ["Bar Chart without gridlines"]
    %w(first second third).each { |label| sheet.add_row [label, rand(24)+1] }
    sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A6", :end_at => "F20") do |chart|
      chart.add_series :data => sheet["B2:B4"], :labels => sheet["A2:A4"], :colors => ["00FF00", "FF0000"]
      chart.valAxis.gridlines = false
      chart.catAxis.gridlines = false
    end
  end
end
#```

##Generating A Pie Chart

#```ruby
if examples.include? :pie_chart
  wb.add_worksheet(:name => "Pie Chart") do |sheet|
    sheet.add_row ["Simple Pie Chart"]
    %w(first second third).each { |label| sheet.add_row [label, rand(24)+1] }
    sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,5], :end_at => [10, 20], :title => "example 3: Pie Chart") do |chart|
      chart.add_series :data => sheet["B2:B4"], :labels => sheet["A2:A4"],  :colors => ['FF0000', '00FF00', '0000FF']
    end
  end
end
#```

##Generating A Line Chart

#```ruby
if examples.include? :line_chart
  wb.add_worksheet(:name => "Line Chart") do |sheet|
    sheet.add_row ["Simple Line Chart"]
    sheet.add_row %w(first second)
    4.times do
      sheet.add_row [ rand(24)+1, rand(24)+1]
    end
    sheet.add_chart(Axlsx::Line3DChart, :title => "Simple 3D Line Chart", :rotX => 30, :rotY => 20) do |chart|
      chart.start_at 0, 5
      chart.end_at 10, 20
      chart.add_series :data => sheet["A3:A6"], :title => sheet["A2"], :color => "0000FF"
      chart.add_series :data => sheet["B3:B6"], :title => sheet["B2"], :color => "FF0000"
      chart.catAxis.title = 'X Axis'
      chart.valAxis.title = 'Y Axis'
    end
    sheet.add_chart(Axlsx::LineChart, :title => "Simple Line Chart", :rotX => 30, :rotY => 20) do |chart|
      chart.start_at 0, 21
      chart.end_at 10, 41
      chart.add_series :data => sheet["A3:A6"], :title => sheet["A2"], :color => "FF0000"
      chart.add_series :data => sheet["B3:B6"], :title => sheet["B2"], :color => "00FF00"
      chart.catAxis.title = 'X Axis'
      chart.valAxis.title = 'Y Axis'
    end

  end
end
#```

##Generating A Scatter Chart

#```ruby
if examples.include? :scatter_chart
  wb.add_worksheet(:name => "Scatter Chart") do |sheet|
    sheet.add_row ["First",  1,  5,  7,  9]
    sheet.add_row ["",       1, 25, 49, 81]
    sheet.add_row ["Second", 5,  2, 14,  9]
    sheet.add_row ["",       5, 10, 15, 20]
    sheet.add_chart(Axlsx::ScatterChart, :title => "example 7: Scatter Chart") do |chart|
      chart.start_at 0, 4
      chart.end_at 10, 19
      chart.add_series :xData => sheet["B1:E1"], :yData => sheet["B2:E2"], :title => sheet["A1"], :color => "FF0000"
      chart.add_series :xData => sheet["B3:E3"], :yData => sheet["B4:E4"], :title => sheet["A3"], :color => "00FF00"
    end
  end
end
#```


##Tables

#```ruby
if examples.include? :tables
  wb.add_worksheet(:name => "Table") do |sheet|
    sheet.add_row ["Build Matrix"]
    sheet.add_row ["Build", "Duration", "Finished", "Rvm"]
    sheet.add_row ["19.1", "1 min 32 sec", "about 10 hours ago", "1.8.7"]
    sheet.add_row ["19.2", "1 min 28 sec", "about 10 hours ago", "1.9.2"]
    sheet.add_row ["19.3", "1 min 35 sec", "about 10 hours ago", "1.9.3"]
    sheet.add_table "A2:D5", :name => 'Build Matrix', :style_info => { :name => "TableStyleMedium23" }
  end
end
#```


##Fit to page printing

#```ruby
if examples.include? :fit_to_page
  wb.add_worksheet(:name => "fit to page") do |sheet|
    sheet.add_row ['this all goes on one page']
    sheet.page_setup.fit_to :width => 1, :height => 1
  end
end
##```


##Hide Gridlines in worksheet

#```ruby
if examples.include? :hide_gridlines
  wb.add_worksheet(:name => "No Gridlines") do |sheet|
    sheet.add_row ["This", "Sheet", "Hides", "Gridlines"]
    sheet.sheet_view.show_grid_lines = false
  end
end
##```

# Repeat printing of header rows.
#```ruby
if examples.include? :repeated_header
  wb.add_worksheet(:name => "repeated header") do |sheet|
    sheet.add_row %w(These Column Header Will Render On Every Printed Sheet)
    200.times { sheet.add_row %w(1 2 3 4 5 6 7 8) }
    wb.add_defined_name("'repeated header'!$1:$1", :local_sheet_id => sheet.index, :name => '_xlnm.Print_Titles')
  end
end

# Defined Names in formula
if examples.include? :defined_name
  wb.add_worksheet(:name => 'defined name') do |sheet|
    sheet.add_row [1, 2, 17, '=FOOBAR']
    wb.add_defined_name("'defined name'!$C1", :local_sheet_id => sheet.index, :name => 'FOOBAR')
  end
end

# Sheet Protection and excluding cells from locking.
if examples.include? :sheet_protection
  unlocked = wb.styles.add_style :locked => false
  wb.add_worksheet(:name => 'Sheet Protection') do |sheet|
    sheet.sheet_protection.password = 'fish'
    sheet.add_row [1, 2 ,3] # These cells will be locked
    sheet.add_row [4, 5, 6], :style => unlocked # these cells will not!
  end
end

##Specify page margins and other options for printing

#```ruby
if examples.include? :printing
  margins = {:left => 3, :right => 3, :top => 1.2, :bottom => 1.2, :header => 0.7, :footer => 0.7}
  setup = {:fit_to_width => 1, :orientation => :landscape, :paper_width => "297mm", :paper_height => "210mm"}
  options = {:grid_lines => true, :headings => true, :horizontal_centered => true}
  wb.add_worksheet(:name => "print margins", :page_margins => margins, :page_setup => setup, :print_options => options) do |sheet|
    sheet.add_row ["this sheet uses customized print settings"]
  end
end
#```

## Add headers and footers to a worksheet
#``` ruby
if examples.include? :header_footer
  header_footer = {:different_first => false, :odd_header => '&L&F : &A&R&D &T', :odd_footer => '&C&Pof&N'}
  wb.add_worksheet(:name => "header footer", :header_footer => header_footer) do |sheet|
    sheet.add_row ["this sheet has a header and a footer"]
  end
end
#```

## Add Comments to your spreadsheet
#``` ruby
if examples.include? :comments
  wb.add_worksheet(:name => 'comments') do |sheet|
    sheet.add_row ['Can we build it?']
    sheet.add_comment :ref => 'A1', :author => 'Bob', :text => 'Yes We Can!'
    sheet.add_comment :ref => 'A2', :author => 'Bob', :text => 'Yes We Can! - but I dont think  you need to know about it!', :visible => false

  end
end

## Frozen/Split panes
## ``` ruby
if examples.include? :panes
  wb.add_worksheet(:name => 'panes') do |sheet|
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
end

if examples.include? :sheet_view
  ws = wb.add_worksheet(:name => 'SheetView - Split')
  ws.sheet_view do |vs|
    vs.pane do |pane|
      pane.active_pane = :top_right
      pane.state = :split
      pane.x_split = 11080
      pane.y_split = 5000
      pane.top_left_cell = 'C44'
    end

    vs.add_selection(:top_left, { :active_cell => 'A2', :sqref => 'A2' })
    vs.add_selection(:top_right, { :active_cell => 'I10', :sqref => 'I10' })
    vs.add_selection(:bottom_left, { :active_cell => 'E55', :sqref => 'E55' })
    vs.add_selection(:bottom_right, { :active_cell => 'I57', :sqref => 'I57' })
  end

  ws = wb.add_worksheet :name => "Sheetview - Frozen"
  ws.sheet_view do |vs|
    vs.pane do |pane|
      pane.state = :frozen
      pane.x_split = 3
      pane.y_split = 4
    end
  end
end

# conditional formatting
#
if examples.include? :conditional_formatting
  percent = wb.styles.add_style(:format_code => "0.00%", :border => Axlsx::STYLE_THIN_BORDER)
  money = wb.styles.add_style(:format_code => '0,000', :border => Axlsx::STYLE_THIN_BORDER)

  # define the style for conditional formatting
  profitable = wb.styles.add_style( :fg_color => "FF428751", :type => :dxf )
  unprofitable = wb.styles.add_style( :fg_color => "FF0000", :type => :dxf )

  wb.add_worksheet(:name => "Conditional Cell Is") do |sheet|

    # Generate 20 rosheet of data
    sheet.add_row ["Previous Year Quarterly Profits (JPY)"]
    sheet.add_row ["Quarter", "Profit", "% of Total"]
    offset = 3
    rosheet = 20
    offset.upto(rosheet + offset) do |i|
      sheet.add_row ["Q#{i}", 10000*((rosheet/2-i) * (rosheet/2-i)), "=100*B#{i}/SUM(B3:B#{rosheet+offset})"], :style=>[nil, money, percent]
    end

    # Apply conditional formatting to range B3:B100 in the worksheet
    sheet.add_conditional_formatting("B3:B100", { :type => :cellIs, :operator => :greaterThan, :formula => "100000", :dxfId => profitable, :priority => 1 })
    # Apply conditional using the between operator; NOTE: supply an array to :formula for between/notBetween
    sheet.add_conditional_formatting("C3:C100", { :type => :cellIs, :operator => :between, :formula => ["0.00%","100.00%"], :dxfId => unprofitable, :priority => 1 })
  end

  wb.add_worksheet(:name => "Conditional Color Scale") do |sheet|
    sheet.add_row ["Previous Year Quarterly Profits (JPY)"]
    sheet.add_row ["Quarter", "Profit", "% of Total"]
    offset = 3
    rosheet = 20
    offset.upto(rosheet + offset) do |i|
      sheet.add_row ["Q#{i}", 10000*((rosheet/2-i) * (rosheet/2-i)), "=100*B#{i}/SUM(B3:B#{rosheet+offset})"], :style=>[nil, money, percent]
    end
    # color scale has two_tone and three_tone class methods to setup the excel defaults (2011)
    # alternatively, you can pass in {:type => [:min, :max, :percent], :val => [whatever], :color =>[Some RGB String] to create a customized color scale object

    color_scale = Axlsx::ColorScale.three_tone
    sheet.add_conditional_formatting("B3:B100", { :type => :colorScale, :operator => :greaterThan, :formula => "100000", :dxfId => profitable, :priority => 1, :color_scale => color_scale })
  end


  wb.add_worksheet(:name => "Conditional Data Bar") do |sheet|
    sheet.add_row ["Previous Year Quarterly Profits (JPY)"]
    sheet.add_row ["Quarter", "Profit", "% of Total"]
    offset = 3
    rows = 20
    offset.upto(rows + offset) do |i|
      sheet.add_row ["Q#{i}", 10000*((rows/2-i) * (rows/2-i)), "=100*B#{i}/SUM(B3:B#{rows+offset})"], :style=>[nil, money, percent]
    end
    # Apply conditional formatting to range B3:B100 in the worksheet
    data_bar = Axlsx::DataBar.new
    sheet.add_conditional_formatting("B3:B100", { :type => :dataBar, :dxfId => profitable, :priority => 1, :data_bar => data_bar })
  end

  wb.add_worksheet(:name => "Conditional Format Icon Set") do |sheet|
    sheet.add_row ["Previous Year Quarterly Profits (JPY)"]
    sheet.add_row ["Quarter", "Profit", "% of Total"]
    offset = 3
    rows = 20
    offset.upto(rows + offset) do |i|
      sheet.add_row ["Q#{i}", 10000*((rows/2-i) * (rows/2-i)), "=100*B#{i}/SUM(B3:B#{rows+offset})"], :style=>[nil, money, percent]
    end
    # Apply conditional formatting to range B3:B100 in the worksheet
    icon_set = Axlsx::IconSet.new
    sheet.add_conditional_formatting("B3:B100", { :type => :iconSet, :dxfId => profitable, :priority => 1, :icon_set => icon_set })
  end
end

##Validate and Serialize

#```ruby
# Serialize directly to file

p.serialize("example.xlsx")

# or

#Serialize to a stream
if examples.include? :streaming
  s = p.to_stream()
  File.open('example_streamed.xlsx', 'w') { |f| f.write(s.read) }
end
#```

##Using Shared Strings

#```ruby
# This is required by Numbers
if examples.include? :shared_strings
  p.use_shared_strings = true
  p.serialize("shared_strings_example.xlsx")
end
#```

#```ruby
if examples.include? :no_autowidth
  p = Axlsx::Package.new
  p.use_autowidth = false
  wb = p.workbook
  wb.add_worksheet(:name => "Manual Widths") do | sheet |
    sheet.add_row ['oh look! no autowidth']
  end
  p.serialize("no-use_autowidth.xlsx")
end
#```



if examples.include? :cached_formula
  p = Axlsx::Package.new
  p.use_shared_strings = true
  wb = p.workbook
  wb.add_worksheet(:name => "cached formula") do | sheet |
    sheet.add_row [1, 2, '=A1+B1'], :formula_values => [nil, nil, 3]
  end
  p.serialize 'cached_formula.xlsx'
end

