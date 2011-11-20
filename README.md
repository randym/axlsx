Axlsx: Office Open XML Spreadsheet Generation
====================================

**IRC**:          [irc.freenode.net / #axlsx](irc://irc.freenode.net/axlsx)    
**Git**:          [http://github.com/randym/axlsx](http://github.com/randym/axlsx)   
**Author**:       Randy Morgan
**Copyright**:    2011    
**License**:      MIT License    
**Latest Version**: 1.0.0
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

Installing
----------

To install Axlsx, use the following command:

    $ gem install axlsx
    
Usage
-----

Generating a workbook with styles and a chart:
	   
	   p = Axlsx::Package.new do |package|
	     package.workbook.add_worksheet do |sheet|
      	       sheet.add_row ["First", "Second", "Third"], :style => Axlsx::STYLE_THIN_BORDER
      	       sheet.add_row [1, 2, 3], :style => Axlsx::STYLE_THIN_BORDER
	       sheet.add_chart(Axlsx::Bar3DChart, :start_at => [0,2], :end_at => [5, 15], :title=>"example 1: Chart") do |chart|
                 chart.add_series :data=>sheet.rows.last.cells, :labels=> sheet.rows.first.cells
	       end
             end
             package.serialize("example1.xlsx")
           end  


### Documentation
This gem is 100% documented with YARD, an exceptional documentation library. To see documentation for this, and all the gems installed on your system use:

  yard server -g


### Specs
This gem has 100% test coverage. To execute tests for this gem, simply run rake in the gem directory.
 
Changelog
---------

- **October.10.11**: 0.1.0 release

Copyright
---------

Axlsx &copy; 2011 by [Randy Morgan](mailto:digial.ipseity@gmail.com). Axlsx is 
licensed under the MIT license. Please see the {file:LICENSE} document for more information.
