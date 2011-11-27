Axlsx: Office Open XML Spreadsheet Generation
====================================

**IRC**:          [irc.freenode.net / #axlsx](irc://irc.freenode.net/axlsx)    
**Git**:          [http://github.com/randym/axlsx](http://github.com/randym/axlsx)   
**Author**:       Randy Morgan   
**Copyright**:    2011      
**License**:      MIT License      
**Latest Version**: 1.0.10.a   
**Ruby Version**: 1.8.7 - 1.9.3   
**Release Date**: November 26th 2011     

Synopsis
--------

Axlsx is an Office Open XML Spreadsheet generator for the Ruby programming language.
With Axlsx you can create worksheets with charts, images, automated column width, customizable styles and full schema validation. Axlsx excels at helping you generate beautiful Office Open XML Spreadsheet documents without having to understand the entire ECMA specification. Check out the README for some examples of how easy it is. Best of all, you can validate your xlsx file before serialization so you know for sure that anything generated is going to load on your client's machine.

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

**6. Support for both 1904 and 1900 epocs configurable in the workbook.

**7. Add jpg, gif and png images to worksheets

**8. Build in mixin with Active record. simply add acts_as_xlsx to you models and they will support to_xlsx

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

Asian Language Support

     p = Axlsx::Package.new
     p.workbook.add_worksheet do |sheet|
       sheet.add_row ["日本語"]
       sheet.add_row ["华语/華語"]
       sheet.add_row ["한국어/조선말"]
     end  
     p.serialize("example8.xlsx")

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
     head = p.workbook.styles.add_style :bg_color => "FF000000", :fg_color=>"FFFFFFFF"
     percent = p.workbook.styles.add_style :num_fmt => 9
     p.workbook.worksheets.first.col_style 2, percent, :row_offset=>1
     p.workbook.worksheets.first.row_style 0, head
     p.serialize("example11.xlsx")


Rails 3

     #ImageMagick port needs to be confirgured --disable-openmp when using with rails3
     #http://stackoverflow.com/questions/2838307/why-is-this-rmagick-call-generating-a-segmentation-fault

     #Add the gem to your Gemfile and bundle install
       gem 'axlsx'

     # A Model that has id, name, title, content, vote, popularity, created_at and updated_at attributes 
     class Post < ActiveRecord::Base
     
       # support for :include, :exclude, :only like ruport has not been implemented yet
       acts_as_axlsx

       def report
         # First generate your package from the model
         # scopes, conditions etc are allowed but must be before the to_xlsx call.
         p = Post.to_xlsx

         # Add in some basic styles and formats
         percent, date, header = nil, nil, nil
         p.workbook.styles do |s|
           percent = s.add_style :num_fmt=>9
           date = s.add_style :format_code => 'yyyy/mm/dd'
           header = s.add_style :bg_color => "FF000000", :fg_color => "FFFFFFFF", :alignment => {:horizontal=>:center}, :sz=>14
         end

         ws = p.workbook.worksheets.first
  
         # Update the row and columns style so they propery format data.
         ws.row_style(0, header)
         ws.col_style(5, percent, :row_offset=>1)
         ws.col_style((6..7), date, :row_offset=>1)
       
         # Add a chart for good measure
         ws.add_chart(Axlsx::Bar3DChart) do |chart|
           chart.title='Post Votes'
           chart.start_at 0, ws.rows.size
           chart.end_at 4, ws.rows.size+20
           chart.add_series :data => ws.cols[4][(1..-1)], :labels=>ws.cols[0][(1..-1)]
         end
    
         p.serialize('posts.xlsx')
         send_file 'posts.xlsx', :type=>"application/xlsx", :x_sendfile=>true
       end
     end


###Documentation

This gem is 100% documented with YARD, an exceptional documentation library. To see documentation for this, and all the gems installed on your system use:

      gem install yard
      yard server -g


###Specs

This gem has 100% test coverage using test/unit. To execute tests for this gem, simply run rake in the gem directory.
 
Changelog
---------
- **October.27.11**: 1.0.10.a release
  - Updating gemspec to loosen up some of the gem requirements in the hope of maintaining compatibility with rails 2
  - Added acts_as_xlsx mixin for rails3 See Examples
  - Added row.style assignation for updating the cell style for an entire row
  - Added col_style method to worksheet upate a the style for a column of cells

- **October.26.11**: 1.0.9 release
  - Updated to support ruby 1.9.3
  - Updated to eliminate all warnings originating in this gem
 
Please see the {file:CHANGELOG.md} document for past release information.


Copyright
---------

Axlsx &copy; 2011 by [Randy Morgan](mailto:digial.ipseity@gmail.com). Axlsx is 
licensed under the MIT license. Please see the {file:LICENSE} document for more information.
