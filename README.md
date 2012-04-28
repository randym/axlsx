Axlsx: Office Open XML Spreadsheet Generation
====================================
[![Build Status](https://secure.travis-ci.org/randym/axlsx.png)](http://travis-ci.org/randym/axlsx/)

**IRC**:[irc.freenode.net / #axlsx](irc://irc.freenode.net/axlsx)

**Git**:[http://github.com/randym/axlsx](http://github.com/randym/axlsx)

**Twitter**: [https://twitter.com/#!/morgan_randy](https://twitter.com/#!/morgan_randy)

**Google Group**: [https://groups.google.com/forum/?fromgroups#!forum/axlsx](https://groups.google.com/forum/?fromgroups#!forum/axlsx)

**Author**:  Randy Morgan

**Copyright**:    2011 - 2012

**License**: MIT License

**Latest Version**: 1.1.3

**Ruby Version**: 1.8.7, 1.9.2, 1.9.3

**JRuby Version**: 1.6.7 1.8 and 1.9 modes

**Rubinius Version**: rubinius 2.0.0dev * lower versions may run, this gem always tests against head.

**Release Date**: April 25th 2012

Synopsis
--------

Axlsx is an Office Open XML Spreadsheet generator for the Ruby programming language.
With Axlsx you can create excel worksheets with charts, images (with links), automated and fixed column widths, customized styles, functions, merged cells, auto filters, file and stream serialization  as well as full schema validation. Axlsx excels at helping you generate beautiful Office Open XML Spreadsheet documents without having to understand the entire ECMA specification.

![Screen 1](https://github.com/randym/axlsx/raw/master/examples/sample.png)

If you are working in rails, or with active record see:
http://github.com/randym/acts_as_xlsx

There are guides for using axlsx and acts_as_xlsx here:
[http://axlsx.blogspot.com](http://axlsx.blogspot.com)


Feature List
------------

**1. Author xlsx documents: Axlsx is made to let you easily and quickly generate professional xlsx based reports that can be validated before serialization.

**2. Generate 3D Pie, Line and Bar Charts: With Axlsx chart generation and management is as easy as a few lines of code. You can build charts based off data in your worksheet or generate charts without any data in your sheet at all.

**3. Custom Styles: With guaranteed document validity, you can style borders, alignment, fills, fonts, and number formats in a single line of code. Those styles can be applied to an entire row, or a single cell anywhere in your workbook.

**4. Automatic type support: Axlsx will automatically determine the type of data you are generating. In this release Float, Integer, String, Date, Time and Boolean types are automatically identified and serialized to your spreadsheet.

**5. Automatic and fixed column widths: Axlsx will automatically determine the appropriate width for your columns based on the content in the worksheet, or use any value you specify for the really funky stuff.

**6. Support for automatically formatted 1904 and 1900 epochs configurable in the workbook.

**7. Add jpg, gif and png images to worksheets with hyperlinks

**8. Reference cells in your worksheet with "A1" and "A1:D4" style references or from the workbook using "Sheet1!A3:B4" style references

**9. Cell level style overrides for default and customized style objects

**10. Support for formulas

**11. Support for cell merging via worksheet.merged_cells

**12. Auto filtering tables with worksheet.auto_filter

**13. Export using shared strings or inline strings so we can inter-op with iWork Numbers (sans charts for now).

**14. Output to file or StringIO

**15. Support for page margins

Installing
----------

To install Axlsx, use the following command:

    $ gem install axlsx

#Usage
------

```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook
```

#A Simple Workbook

```ruby
wb.add_worksheet(:name => "Basic Worksheet") do |sheet|
  sheet.add_row ["First Column", "Second", "Third"]
  sheet.add_row [1, 2, 3]
end
```

#Using Custom Styles

```ruby
wb.styles do |s|
  black_cell = s.add_style :bg_color => "00", :fg_color => "FF", :sz => 14, :alignment => { :horizontal=> :center }
  blue_cell =  s.add_style  :bg_color => "0000FF", :fg_color => "FF", :sz => 20, :alignment => { :horizontal=> :center }
  wb.add_worksheet(:name => "Custom Styles") do |sheet|
    sheet.add_row ["Text Autowidth", "Second", "Third"], :style => [black_cell, blue_cell, black_cell]
    sheet.add_row [1, 2, 3], :style => Axlsx::STYLE_THIN_BORDER
  end
end
```

#Using Custom Border Styles

```ruby
wb.styles do |s|
  red_border =  s.add_style :border => { :style => :thin, :color =>"FFFF0000" }
  blue_border =  s.add_style :border => { :style => :thin, :color =>"FF0000FF" }

  wb.add_worksheet(:name => "Custom Borders") do |sheet|
    sheet.add_row ["wrap", "me", "Up in Red"], :style => red_border
    sheet.add_row [1, 2, 3], :style => blue_border
  end
end
```


##Using Custom Formatting and date1904

```ruby
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
```


##Add an Image

```ruby
wb.add_worksheet(:name => "Images") do |sheet|
  img = File.expand_path('../image1.jpeg', __FILE__)
  sheet.add_image(:image_src => img, :noSelect => true, :noMove => true) do |image|
    image.width=720
    image.height=666
    image.start_at 2, 2
  end
end
```

##Add an Image with a hyperlink

```ruby
wb.add_worksheet(:name => "Image with Hyperlink") do |sheet|
  img = File.expand_path('../image1.jpeg', __FILE__)
  sheet.add_image(:image_src => img, :noSelect => true, :noMove => true, :hyperlink=>"http://axlsx.blogspot.com") do |image|
    image.width=720
    image.height=666
    image.hyperlink.tooltip = "Labeled Link"
    image.start_at 2, 2
  end
end
```

##Asian Language Support

```ruby
wb.add_worksheet(:name => "日本語でのシート名") do |sheet|
  sheet.add_row ["日本語"]
  sheet.add_row ["华语/華語"]
  sheet.add_row ["한국어/조선말"]
end
```

##Styling Columns

```ruby
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
```

##Hiding Columns

```ruby
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
```

##Styling Rows

```ruby
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
```

##Styling Cell Overrides

```ruby
wb.add_worksheet(:name => "Cell Level Style Overrides") do |sheet|
  # cell level style overides when adding cells
  sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4'], :sz => 16
  sheet.add_row [1, 2, 3, "=SUM(A2:C2)"]
  # cell level style overrides via sheet range
  sheet["A1:D1"].each { |c| c.color = "FF0000"}
  sheet['A1:D2'].each { |c| c.style = Axlsx::STYLE_THIN_BORDER }
end
```

##Using formula

```ruby
wb.add_worksheet(:name => "Using Formulas") do |sheet|
  sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4']
  sheet.add_row [1, 2, 3, "=SUM(A2:C2)"]
end
```

##Automatic cell types

```ruby
wb.add_worksheet(:name => "Automatic cell types") do |sheet|
  sheet.add_row ["Date", "Time", "String", "Boolean", "Float", "Integer"]
  sheet.add_row [Date.today, Time.now, "value", true, 0.1, 1]
end
```

##Merging Cells.

```ruby
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
```

##Generating A Bar Chart

```ruby
wb.add_worksheet(:name => "Bar Chart") do |sheet|
  sheet.add_row ["A Simple Bar Chart"]
  sheet.add_row ["First", "Second", "Third"]
  sheet.add_row [1, 2, 3]
  sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A4", :end_at => "F17") do |chart|
    chart.add_series :data => sheet["A3:C3"], :labels => sheet["A2:C2"], :title => sheet["A1"]
    chart.valAxis.label_rotation = -45
    chart.catAxis.label_rotation = 45
  end
end
```

##Hide Gridlines in chart

```ruby
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
```

##Generating A Pie Chart

```ruby
wb.add_worksheet(:name => "Pie Chart") do |sheet|
  sheet.add_row ["First", "Second", "Third", "Fourth"]
  sheet.add_row [1, 2, 3, "=PRODUCT(A2:C2)"]
  sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,2], :end_at => [5, 15], :title => "example 3: Pie Chart") do |chart|
    chart.add_series :data => sheet["A2:D2"], :labels => sheet["A1:D1"]
  end
end
```

##Data over time

```ruby
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
```

##Generating A Line Chart

```ruby
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
```

##Generating A Scatter Chart

```ruby
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
```

##Auto Filter

```ruby
wb.add_worksheet(:name => "Auto Filter") do |sheet|
  sheet.add_row ["Build Matrix"]
  sheet.add_row ["Build", "Duration", "Finished", "Rvm"]
  sheet.add_row ["19.1", "1 min 32 sec", "about 10 hours ago", "1.8.7"]
  sheet.add_row ["19.2", "1 min 28 sec", "about 10 hours ago", "1.9.2"]
  sheet.add_row ["19.3", "1 min 35 sec", "about 10 hours ago", "1.9.3"]
  sheet.auto_filter = "A2:D5"
end
```

##Tables

```ruby
wb.add_worksheet(:name => "Table") do |sheet|
  sheet.add_row ["Build Matrix"]
  sheet.add_row ["Build", "Duration", "Finished", "Rvm"]
  sheet.add_row ["19.1", "1 min 32 sec", "about 10 hours ago", "1.8.7"]
  sheet.add_row ["19.2", "1 min 28 sec", "about 10 hours ago", "1.9.2"]
  sheet.add_row ["19.3", "1 min 35 sec", "about 10 hours ago", "1.9.3"]
  sheet.add_table "A2:D5", :name => 'Build Matrix'
end
```

##Specifying Column Widths

```ruby
wb.add_worksheet(:name => "custom column widths") do |sheet|
  sheet.add_row ["I use autowidth and am very wide", "I use a custom width and am narrow"]
  sheet.add_row ['abcdefg', 'This is a very long text and should flow into the right cell', nil, 'xxx' ]
  sheet.column_widths nil, 3, 5, nil
end
```

##Fit to page printing

```ruby
wb.add_worksheet(:name => "fit to page") do |sheet|
  sheet.add_row ['this all goes on one page']
  sheet.fit_to_page = true
end
```

##Hide Gridlines in worksheet

```ruby
wb.add_worksheet(:name => "No Gridlines") do |sheet|
  sheet.add_row ["This", "Sheet", "Hides", "Gridlines"]
  sheet.show_gridlines = false
end
```

##Specify Page Margins for printing

```ruby
margins = {:left => 3, :right => 3, :top => 1.2, :bottom => 1.2, :header => 0.7, :footer => 0.7}
wb.add_worksheet(:name => "print margins", :page_margins => margins) do |sheet|
  sheet.add_row ["this sheet uses customized page margins for printing"]
end
```

##Validate and Serialize

```ruby
p.serialize("example.xlsx")

s = p.to_stream()
File.open('example_streamed.xlsx', 'w') { |f| f.write(s.read) }
```

##Using Shared Strings

```ruby
p.use_shared_strings = true
p.serialize("shared_strings_example.xlsx")
```

##Disabling Autowidth

```ruby
p = Axlsx::Package.new
p.use_autowidth = false
wb = p.workbook
wb.add_worksheet(:name => "Manual Widths") do | sheet |
  sheet.add_row ['oh look! no autowidth']
end
p.validate.each { |e| puts e.message }
p.serialize("no-use_autowidth.xlsx")
```

#Documentation
--------------
This gem is 100% documented with YARD, an exceptional documentation library. To see documentation for this, and all the gems installed on your system use:

```bash
gem install yard
yard server -g
```
#Specs
------
This gem has 100% test coverage using test/unit. To execute tests for this gem, simply run rake in the gem directory.

#Change log
---------
- ** April.??.12:**: 1.1.4 release
   - bugfix in val_axis_data to properly serialize value axis data. Excel does not mind as it reads from the sheet, but nokogiri has a fit if the elements are empty.
   - Added support for specifying the color of data series in charts.
   - bugfix using add_cell on row mismanaged calls to update_column_info.

- ** April.25.12:**: 1.1.3 release
   - Primarily because I am stupid.....Updates to readme to properly report version, add in missing docs and restructure example directory.

- ** April.25.12:**: 1.1.2 release
   - Conditional Formatting completely implemented.
   - refactoring / documentation for Style#add_style
   - added in label rotation for chart axis labels
   - bugfix to properly assign style and type info to cells when only partial information is provided in the types/style option

- ** April.18.12**: 1.1.1 release
   - bugfix for autowidth calculations across multiple rows
   - bugfix for dimension calculations with nil cells.
   - REMOVED RMAGICK dependency WOOT!
   - Update readme to show screenshot of gem output.
   - Cleanup benchmark and add benchmark rake task

- ** April.3.12**: 1.1.0 release
   - bugfix patch name_to_indecies to properly handle extended ranges.
   - bugfix properly serialize chart title.
   - lower rake minimum requirement for 1.8.7 apps that don't want to move on to 0.9 NOTE this will be reverted for 2.0.0 with workbook parsing!
   - Added Fit to Page printing
   - added support for turning off gridlines in charts.
   - added support for turning off gridlines in worksheet.
   - bugfix some apps like libraoffice require apply[x] attributes to be true. applyAlignment is now properly set.
   - added option use_autowidth. When this is false RMagick will not be loaded or used in the stack. However it is still a requirement in the gem.
   - added border style specification to styles#add_style. See the example in the readme.
   - Support for tables added in - Note: Pre 2011 versions of Mac office do not support this feature and will warn.
   - Support for splatter charts added
   - Major (like 7x faster!) performance updates.
   - Gem now supports for JRuby 1.6.7, as well as experimental support for Rubinius

Please see the {file:CHANGELOG.md} document for past release information.

#Thanks!
--------
[ochko](https://github.com/ochko) - for performance fixes, kicking the crap out of axlsx and helping to maintain my general sanity.

[kleine2](https://github.com/kleine2) - for generously donating in return for the image hyperlink feature.

[ffmike](https://github.com/ffmike) - for knocking down an over restrictive i18n dependency, massive patience and great communication skills.

[JonathanTron](https://github.com/JonathanTron) - for giving the gem some style, and making sure it applies.

[JosephHalter](https://github.com/JosephHalter) - for making sure we arrive at the right time on the right date.

[noniq](https://github.com/noniq) - for keeping true to the gem's style, and making sure what we put on paper does not get marginalized.

[jurriaan](https://github.com/jurriaan) - for showing there is more than one way to skin a cat, and work with rows while you are at it.

[joekain](https://github.com/joekain) - for keeping our references working even in the double digits!

[moskrin](https://github.com/moskrin) - for keeping border creation on the edge.

[scpike](https://github.com/scpike) - for keeping numbers fixed even when they are rational and a super clean implementation of conditional formatting.

#Copyright and License
----------

Axlsx &copy; 2011-2012 by [Randy Morgan](mailto:digial.ipseity@gmail.com). Axlsx is
licensed under the MIT license. Please see the LICENSE document for more information.
