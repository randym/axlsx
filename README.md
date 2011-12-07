Axlsx: Office Open XML Spreadsheet Generation
====================================
[![Build Status](https://secure.travis-ci.org/randym/axlsx.png)](http://travis-ci.org/randym/axlsx/)

**IRC**:          [irc.freenode.net / #axlsx](irc://irc.freenode.net/axlsx)    
**Git**:          [http://github.com/randym/axlsx](http://github.com/randym/axlsx)   
**Author**:       Randy Morgan   
**Copyright**:    2011      
**License**:      MIT License      
**Latest Version**: 1.0.12   
**Ruby Version**: 1.8.7, 1.9.2, 1.9.3 

**Release Date**: December 7th 2011     

Synopsis
--------

Axlsx is an Office Open XML Spreadsheet generator for the Ruby programming language.
With Axlsx you can create worksheets with charts, images, automated column width, customizable styles and full schema validation. Axlsx excels at helping you generate beautiful Office Open XML Spreadsheet documents without having to understand the entire ECMA specification. Check out the README for some examples of how easy it is. Best of all, you can validate your xlsx file before serialization so you know for sure that anything generated is going to load on your client's machine.

If you are working in rails, or with active record see:
http://github.com/randym/acts_as_xlsx 

Help Wanted
-----------

I'd really like to get rid of the depenency on RMagick in this gem. RMagic is being used to calculate the column widths in a worksheet based on the content the user specified. If there happens to be anyone out there with the background and c skills to write an extenstion that can determine the width of a single character rendered with a specific font at a specific font size please give me a shout.

Feature List
------------
                                                                              
**1. Author xlsx documents: Axlsx is made to let you easily and quickly generate profesional xlsx based reports that can be validated before serialiation.

**2. Generate 3D Pie, Line and Bar Charts: With Axlsx chart generation and management is as easy as a few lines of code. You can build charts based off data in your worksheet or generate charts without any data in your sheet at all.
                                                                              
**3. Custom Styles: With guaranteed document validity, you can style borders, alignment, fills, fonts, and number formats in a single line of code. Those styles can be applied to an entire row, or a single cell anywhere in your workbook.

**4. Automatic type support: Axlsx will automatically determine the type of data you are generating. In this release Float, Integer, String and Time types are automatically identified and serialized to your spreadsheet.

**5. Automatic column widths: Axlsx will automatically determine the appropriate width for your columns based on the content in the worksheet.

**6. Support for automatically formatted 1904 and 1900 epocs configurable in the workbook.

**7. Add jpg, gif and png images to worksheets

**8. Refernce cells in your worksheet with "A1" and "A1:D4" style references or from the workbook using "Sheett1!A3:B4" style references

**9. Cell level style overrides for default and customized style object

**10. Support for formulas

Installing
----------

To install Axlsx, use the following command:

    $ gem install axlsx
    
Usage
-----

###Examples

     require 'rubygems'
     require 'axlsx'


Validation

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["First", "Second", "Third"]
       sheet.add_row [1, 2, 3]
     end

     p.validate.each do |error|
       puts error.inspect
     end

A Simple Workbook

      p = Axlsx::Package.new
      p.workbook.add_worksheet do |sheet|
        sheet.add_row ["First Column", "Second Column", "Third Column"]
        sheet.add_row [1, 2, Time.now]
      end
      p.serialize("example1.xlsx")

Generating A Bar Chart

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["A Simple Bar Chart"]
       sheet.add_row ["First", "Second", "Third"]
       sheet.add_row [1, 2, 3]
       sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A4", :end_at => "F17", :title=>sheet["A1"]) do |chart|
         chart.add_series :data => sheet["A3:C3"], :labels => sheet["A2:C2"]
       end
     end  
     p.serialize("example2.xlsx")

Generating A Pie Chart

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["First", "Second", "Third", "Fourth"]
       sheet.add_row [1, 2, 3, "=PRODUCT(A2:C2)"]
       sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0, 2], :end_at => [5, 25], :title=>"example 3: Pie Chart") do |chart|
         chart.add_series :data => sheet["A2:D2"], :labels => sheet["A1:D1"]
       end
     end  
     p.serialize("example3.xlsx")

Generating A Line Chart

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["First", 1, 5, 7, 9]
       sheet.add_row ["Second", 5, 2, 14, 9]
       sheet.add_chart(Axlsx::Line3DChart, :title=>"example 6: Line Chart") do |chart|
         chart.start_at 0, 2
         chart.end_at 10, 15
         chart.add_series :data=>sheet["B1:E1"], :title=> sheet["A1"]
         chart.add_series :data=>sheet["B2:E2"], :title=> sheet["A2"]
       end
       
     end  
     p.serialize("example4.xlsx")

Using Custom Styles

     p = Axlsx::Package.new
     wb = p.workbook
     black_cell = wb.styles.add_style :bg_color => "00", :fg_color => "FF", :sz=>14, :alignment => { :horizontal=> :center }
     blue_cell = wb.styles.add_style  :bg_color => "0000FF", :fg_color => "FF", :sz=>14, :alignment => { :horizontal=> :center }
     wb.add_worksheet do |sheet|
       sheet.add_row ["Text Autowidth", "Second", "Third"], :style => [black_cell, blue_cell, black_cell]
       sheet.add_row [1, 2, 3], :style => Axlsx::STYLE_THIN_BORDER
     end
     p.serialize("example5.xlsx")

Using Custom Formatting and date1904

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
     p.serialize("example6.xlsx")



Add an Image

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       img = File.expand_path('examples/image1.jpeg') 
       sheet.add_image(:image_src => img, :noSelect=>true, :noMove=>true) do |image|
         image.width=720
         image.height=666
         image.start_at 2, 2
       end
     end  
     p.serialize("example8.xlsx")

Asian Language Support

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["日本語"]
       sheet.add_row ["华语/華語"]
       sheet.add_row ["한국어/조선말"]
     end  
     p.serialize("example9.xlsx")

Styling Columns

     p = Axlsx::Package.new
     percent = p.workbook.styles.add_style :num_fmt => 9
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4']
       sheet.add_row [1, 2, 0.3, 4]
       sheet.add_row [1, 2, 0.2, 4]
       sheet.add_row [1, 2, 0.1, 4]
     end
     p.workbook.worksheets.first.col_style 2, percent, :row_offset=>1
     p.serialize("example10.xlsx")

Styling Rows

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4']
       sheet.add_row [1, 2, 0.3, 4]
       sheet.add_row [1, 2, 0.2, 4]
       sheet.add_row [1, 2, 0.1, 4]
     end
     head = p.workbook.styles.add_style :bg_color => "00", :fg_color=>"FF"
     percent = p.workbook.styles.add_style :num_fmt => 9
     p.workbook.worksheets.first.col_style 2, percent, :row_offset=>1
     p.workbook.worksheets.first.row_style 0, head
     p.serialize("example11.xlsx")

Using formula

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4']
       sheet.add_row [1, 2, 3, "=SUM(A2:C2)"]
     end
     p.serialize("example12.xlsx")

Using cell specific styling and range / name based access

     p = Axlsx::Package.new
     p.workbook.add_worksheet(:name=>'My Worksheet') do |sheet|
         # cell level style overides when adding cells
         sheet.add_row ['col 1', 'col 2', 'col 3', 'col 4'], :sz => 16
         sheet.add_row [1, 2, 3, "=SUM(A2:C2)"]
         # cell level style overrides via sheet range
         sheet["A1:D1"].each { |c| c.color = "FF0000"}
     end     
     p.workbook['My Worksheet!A1:D2'].each { |c| c.style = Axlsx::STYLE_THIN_BORDER }
     p.serialize("example13.xlsx")



###Documentation

This gem is 100% documented with YARD, an exceptional documentation library. To see documentation for this, and all the gems installed on your system use:

      gem install yard
      yard server -g


###Specs

This gem has 100% test coverage using test/unit. To execute tests for this gem, simply run rake in the gem directory.
 
Changelog
---------
- **December.7.11**: 1.0.12 release
  - changed dependency from 'zip' gem to 'rubyzip' and added conditional code to force binary encoding to resolve issue with excel 2011
  - Patched bug in app.xml that would ignore user specified properties.
- **December.5.11**: 1.0.11 release
  - Added [] methods to worksheet and workbook to provide name based access to cells.
  - Added support for functions as cell values
  - Updated color creation so that two character shorthand values can be used like 'FF' for 'FFFFFFFF' or 'D8' for 'FFD8D8D8'
  - Examples for all this fun stuff added to the readme
  - Clean up and support for 1.9.2 and travis integration
  - Added support for string based cell referencing to chart start_at and end_at. That means you can now use :start_at=>"A1" when using worksheet.add_chart, or chart.start_at ="A1" in addition to passing a cell or the x, y coordinates.
 
Please see the {file:CHANGELOG.md} document for past release information.


Copyright
---------

Axlsx &copy; 2011 by [Randy Morgan](mailto:digial.ipseity@gmail.com). Axlsx is 
licensed under the MIT license. Please see the {file:LICENSE} document for more information.
