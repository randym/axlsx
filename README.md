Axlsx: Office Open XML Spreadsheet Generation
====================================

**IRC**:          [irc.freenode.net / #axlsx](irc://irc.freenode.net/axlsx)    
**Git**:          [http://github.com/randym/axlsx](http://github.com/randym/axlsx)   
**Author**:       Randy Morgan
**Copyright**:    2011    
**License**:      MIT License    
**Latest Version**: 1.0.2
**Release Date**: November 20th 2011    

Synopsis
--------

Axlsx is an Office Open XML Spreadsheet generator for the Ruby programming language.
It enables the you to generate 100% valid xlsx files that include customised styling 3D pie, bar and line charts. Below is a summary of salient features.

Feature List
------------
                                                                              
**1. Author xlsx documents: Axlsx is made to let you easily and quickly generate profesional xlsx based reports that can be validated before serialiation.

**2. Generate 3D Pie and Bar Charts: With Axlsx chart generation and management is as easy as a few lines of code. You can build charts based off data in your worksheet or generate charts without any data in your sheet at all.
                                                                              
**3. Custom Styles: With guaranteed document validity, you can style borders, alignment, fills, fonts, and number formats in a single line of code. Those styles can be applied to an entire row, or a single cell anywhere in your workbook.

**4. Automatic type support: Axlsx will automatically determine the type of data you are generating. In this release Float, Integer, String and Time types are automatically identified and serialized to your spreadsheet.

**5. Automatic column widths: Axlsx will automatically determine the appropriate width for your columns based on the content in the worksheet.

**6. Support for both 1904 and 1900 epocs configurable in the workbook.


Installing
----------

To install Axlsx, use the following command:

    $ gem install axlsx
    
Usage
-----
Simple Workbook

     p = Axlsx::Package.new do |package|
       package.workbook.add_worksheet do |sheet|
         sheet.add_row ["First", "Second", "Third"]
         sheet.add_row [1, 2, 3]
       end
       package.serialize("example1.xlsx")
     end  

Generating A Bar Chart
	   
     p = Axlsx::Package.new do |package|
       package.workbook.add_worksheet do |sheet|
         sheet.add_row ["First", "Second", "Third"]
         sheet.add_row [1, 2, 3]
         sheet.add_chart(Axlsx::Bar3DChart, :start_at => [0,2], :end_at => [5, 15], :title=>"example 1: Chart") do |chart|
           chart.add_series :data=>sheet.rows.last.cells, :labels=> sheet.rows.first.cells
         end
       end
       package.serialize("example1.xlsx")
     end  

Generating A Pie Chart

     p = Axlsx::Package.new do |package|
       package.workbook.add_worksheet do |sheet|
         sheet.add_row ["First", "Second", "Third"]
         sheet.add_row [1, 2, 3]
         sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,2], :end_at => [5, 15], :title=>"example 2: Pie Chart") do |chart|
           chart.add_series :data=>sheet.rows.last.cells, :labels=> sheet.rows.first.cells
         end
       end
       package.serialize("example3.xlsx")
     end  

Using Custom Styles

     p = Axlsx::Package.new do |package|
       style_options = { :bg_color => "FF000000", :fg_color => "FFFFFFFF", :sz=>14, :alignment => { :horizontal=> :center } }
       header_style = package.workbook.styles.add_style style_options
       package.workbook.add_worksheet do |sheet|
         sheet.add_row ["Text Autowidth", "Second", "Third"], :style => header_style
         sheet.add_row [1, 2, 3], :style => Axlsx::STYLE_THIN_BORDER
       end
       package.serialize("example3.xlsx")
     end  

Cell Specifc Styles

     p = Axlsx::Package.new do |package|

       black_cell_spec = { :bg_color => "FF000000", :fg_color => "FFFFFFFF", :sz=>14, :alignment => { :horizontal=> :center } }
       blue_cell_spec = { :bg_color => "FF0000FF", :fg_color => "FFFFFFFF", :sz=>14, :alignment => { :horizontal=> :center } }

       black_cell = package.workbook.styles.add_style black_cell_spec
       blue_cell = package.workbook.styles.add_style blue_cell_spec

       # date1904 support. uncomment the line below if you are working on a mac.
       # package.workbook.date1904 = true 

       package.workbook.add_worksheet do |sheet|
         sheet.add_row ["Text Autowidth", "Second", "Third"], :style => [black_cell, blue_cell, black_cell]
         sheet.add_row [1, 2, 3], :style => Axlsx::STYLE_THIN_BORDER
       end
       package.serialize("example3.xlsx")
     end  

Number and Date formatting

     p = Axlsx::Package.new do |package|
       date = package.workbook.styles.add_style :format_code=>"yyyy-mm-dd", :border => Axlsx::STYLE_THIN_BORDER
       padded = package.workbook.styles.add_style :format_code=>"00#", :border => Axlsx::STYLE_THIN_BORDER
       percent = package.workbook.styles.add_style :format_code=>"0%", :border => Axlsx::STYLE_THIN_BORDER

       package.workbook.add_worksheet do |sheet|
         sheet.add_row
         sheet.add_row ["Custom Formatted Date", "Percent Formatted Float", "Padded Numbers"], :style => Axlsx::STYLE_THIN_BORDER
         sheet.add_row [Time.now, 0.2, 32], :style => [date, percent, padded]
       end
       package.serialize("example5.xlsx")
     end  

### Documentation
This gem is 100% documented with YARD, an exceptional documentation library. To see documentation for this, and all the gems installed on your system use:

  yard server -g


### Specs
This gem has 100% test coverage. To execute tests for this gem, simply run rake in the gem directory.
 
Changelog
---------

- **October.20.11**: 0.1.0 release
- **October.21.11**: 1.0.3 release
  - Updated documentation
  - altered package to accept a filename string for serialization instead of a File object.
  - Updated specs to conform

Copyright
---------

Axlsx &copy; 2011 by [Randy Morgan](mailto:digial.ipseity@gmail.com). Axlsx is 
licensed under the MIT license. Please see the {file:LICENSE} document for more information.
