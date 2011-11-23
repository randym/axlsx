Axlsx: Office Open XML Spreadsheet Generation
====================================

**IRC**:          [irc.freenode.net / #axlsx](irc://irc.freenode.net/axlsx)    
**Git**:          [http://github.com/randym/axlsx](http://github.com/randym/axlsx)   
**Author**:       Randy Morgan   
**Copyright**:    2011      
**License**:      MIT License      
**Latest Version**: 1.0.8   
**Ruby Version**: 1.8.7   
**Release Date**: November 23nd 2011     

Synopsis
--------

Axlsx is an Office Open XML Spreadsheet generator for the Ruby programming language.
It enables the you to generate 100% valid xlsx files that include customised styling, images, 3D pie, bar and line charts and auto-width columns. Below is a summary of salient features.

Feature List
------------
                                                                              
**1. Author xlsx documents: Axlsx is made to let you easily and quickly generate profesional xlsx based reports that can be validated before serialiation.

**2. Generate 3D Pie, Line and Bar Charts: With Axlsx chart generation and management is as easy as a few lines of code. You can build charts based off data in your worksheet or generate charts without any data in your sheet at all.
                                                                              
**3. Custom Styles: With guaranteed document validity, you can style borders, alignment, fills, fonts, and number formats in a single line of code. Those styles can be applied to an entire row, or a single cell anywhere in your workbook.

**4. Automatic type support: Axlsx will automatically determine the type of data you are generating. In this release Float, Integer, String and Time types are automatically identified and serialized to your spreadsheet.

**5. Automatic column widths: Axlsx will automatically determine the appropriate width for your columns based on the content in the worksheet.

**6. Support for both 1904 and 1900 epocs configurable in the workbook.

**7. Add jpg, gif and png images to worksheets

Installing
----------

To install Axlsx, use the following command:

    $ gem install axlsx
    
Usage
-----

###Examples

     require 'rubygems'
     require 'axlsx'

A Simple Workbook

      p = Axlsx::Package.new
      p.workbook.add_worksheet do |sheet|
        sheet.add_row ["First", "Second", "Third"]
        sheet.add_row [1, 2, 3]
      end
      p.serialize("example1.xlsx")

Generating A Bar Chart

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["First", "Second", "Third"]
       sheet.add_row [1, 2, 3]
       sheet.add_chart(Axlsx::Bar3DChart, :start_at => [0,2], :end_at => [5, 15], :title=>"example 2: Chart") do |chart|
         chart.add_series :data=>sheet.rows.last.cells, :labels=> sheet.rows.first.cells
       end
     end  
     p.serialize("example2.xlsx")

Generating A Pie Chart

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["First", "Second", "Third"]
       sheet.add_row [1, 2, 3]
       sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,2], :end_at => [5, 15], :title=>"example 3: Pie Chart") do |chart|
         chart.add_series :data=>sheet.rows.last.cells, :labels=> sheet.rows.first.cells
       end
     end  
     p.serialize("example3.xlsx")

Using Custom Styles

     p = Axlsx::Package.new
     wb = p.workbook
     black_cell = wb.styles.add_style :bg_color => "FF000000", :fg_color => "FFFFFFFF", :sz=>14, :alignment => { :horizontal=> :center }
     blue_cell = wb.styles.add_style  :bg_color => "FF0000FF", :fg_color => "FFFFFFFF", :sz=>14, :alignment => { :horizontal=> :center }
     wb.add_worksheet do |sheet|
       sheet.add_row ["Text Autowidth", "Second", "Third"], :style => [black_cell, blue_cell, black_cell]
       sheet.add_row [1, 2, 3], :style => Axlsx::STYLE_THIN_BORDER
     end
     p.serialize("example4.xlsx")

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
     p.serialize("example5.xlsx")

Validation

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["First", "Second", "Third"]
       sheet.add_row [1, 2, 3]
     end

     p.validate.each do |error|
       puts error.inspect
     end

Generating A Line Chart

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["First", 1, 5, 7, 9]
       sheet.add_row ["Second", 5, 2, 14, 9]
       sheet.add_chart(Axlsx::Line3DChart, :title=>"example 6: Line Chart") do |chart|
         chart.start_at 0, 2
         chart.end_at 10, 15
         chart.add_series :data=>sheet.rows.first.cells[(1..-1)], :title=> sheet.rows.first.cells.first
         chart.add_series :data=>sheet.rows.last.cells[(1..-1)], :title=> sheet.rows.last.cells.first
       end
       
     end  
     p.serialize("example6.xlsx")

Adding an Image

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_image(:image_src => (File.dirname(__FILE__) + "/image1.png")) do |image|
         image.width=720
         image.height=666
         image.start_at 2, 2
       end
     end  
     p.serialize("example7.xlsx")

###Documentation

This gem is 100% documented with YARD, an exceptional documentation library. To see documentation for this, and all the gems installed on your system use:

      gem install yard
      yard server -g


###Specs

This gem has 100% test coverage using test/unit. To execute tests for this gem, simply run rake in the gem directory.
 
Changelog
---------
- **October.23.11**: 1.0.8 release
  - Added support for images (jpg, gif, png) in worksheets.

- **October.23.11**: 1.0.7 released
  - Added support for 3D options when creating a new chart. This lets you set the persective, rotation and other 3D attributes when using worksheet.add_chart
  - Updated serialization write test to verify write permissions and warn if it cannot run the test due to permission restrcitions.
  - updated rake to include build, genoc and deploy tasks.
  - rebuilt documentation.
  - moved version constant to its own file
  - fixed bug in SerAxis that was requiring tickLblSkip and tickMarkSkip to be boolean. Should be unsigned int.
  - Review and improve docs
  - rebuild of anchor positioning to remove some spagetti code. Chart now supports a start_at and end_at method that accept an arrar for col/row positioning. See example6 for an example. You can still pass :start_at and :end_at options to worksheet.add_chart.
  - Refactored cat and val axis data to keep series serialization a bit more DRY
 
Please see the {file:CHANGELOG.md} document for past release information.

Help Wanted
-----------

I'd really like to get rid of the depenency on RMagick in this gem. RMagic is being used to calculate the column widths in a worksheet based on the content the user specified. If there happens to be anyone out there with the background and skill set to write an extenstion that can determine the width of a singel character render with a specific font at a specific size please give me a shout.

Copyright
---------

Axlsx &copy; 2011 by [Randy Morgan](mailto:digial.ipseity@gmail.com). Axlsx is 
licensed under the MIT license. Please see the {file:LICENSE} document for more information.
