Axlsx: Office Open XML Spreadsheet Generation
====================================
[![Build Status](https://secure.travis-ci.org/randym/axlsx.png?branch=master)](http://travis-ci.org/randym/axlsx/)

If you are using axlsx for comercial purposes, or just want to show your
appreciation for the gem, please don't hesitate to make a donation.

[![Click here to lend your support to: axlsx and make a donation at www.pledgie.com !](http://www.pledgie.com/campaigns/17814.png?skin_name=chrome)](http://www.pledgie.com/campaigns/17814)

**IRC**:[irc.freenode.net / #axlsx](irc://irc.freenode.net/axlsx)

**Git**:[http://github.com/randym/axlsx](http://github.com/randym/axlsx)

**Twitter**: [https://twitter.com/#!/morgan_randy](https://twitter.com/#!/morgan_randy)

**Google Group**: [https://groups.google.com/forum/?fromgroups#!forum/axlsx](https://groups.google.com/forum/?fromgroups#!forum/axlsx)

**Author**: Randy Morgan

**Copyright**: 2011 - 2013

**License**: MIT License

**Latest Version**: 2.0.1

**Ruby Version**: 1.9.2, 1.9.3, 2.0.0

**JRuby Version**: 1.9 modes

**Rubinius Version**: rubinius 2.0.0dev * lower versions may run, this gem always tests against head.

**Release Date**: September 12th 2013

If you are working in rails, or with active record see:
[acts_as_xlsx](http://github.com/randym/acts_as_xlsx)

acts_as_xlsx is a simple ActiveRecord mixin that lets you generate a workbook with:

```ruby
Posts.where(created_at > Time.now-30.days).to_xlsx
```

** and **

* http://github.com/straydogstudio/axlsx_rails
Axlsx_Rails provides an Axlsx renderer so you can move all your spreadsheet code from your controller into view files. Partials are supported so you can organize any code into reusable chunks (e.g. cover sheets, common styling, etc.) You can use it with acts_as_xlsx, placing the to_xlsx call in a view and add ':package => xlsx_package' to the parameter list. Now you can keep your controllers thin!

There are guides for using axlsx and acts_as_xlsx here:
[http://axlsx.blog.randym.net](http://axlsx.blog.randym.net)

If you are working with ActiveAdmin see:

[activeadmin_axlsx](http://github.com/randym/activeadmin_axlsx)

It provies a plugin and dsl for generating downloadable reports.

The examples directory contains a number of more specific examples as
well.

Synopsis
--------

Axlsx is an Office Open XML Spreadsheet generator for the Ruby programming language.
With Axlsx you can create excel worksheets with charts, images (with links), automated and fixed column widths, customized styles, functions, tables, conditional formatting, print options, comments, merged cells, auto filters, file and stream serialization  as well as full schema validation. Axlsx excels at helping you generate beautiful Office Open XML Spreadsheet documents without having to understand the entire ECMA specification.

![Screen 1](https://github.com/randym/axlsx/raw/master/examples/sample.png)



Feature List
------------

1. Author xlsx documents: Axlsx is made to let you easily and quickly generate professional xlsx based reports that can be validated before serialization.

2. Generate 3D Pie, Line, Scatter and Bar Charts: With Axlsx chart generation and management is as easy as a few lines of code. You can build charts based off data in your worksheet or generate charts without any data in your sheet at all. Customize gridlines, label rotation and series colors as well.

3. Custom Styles: With guaranteed document validity, you can style borders, alignment, fills, fonts, and number formats in a single line of code. Those styles can be applied to an entire row, or a single cell anywhere in your workbook.

4. Automatic type support: Axlsx will automatically determine the type of data you are generating. In this release Float, Integer, String, Date, Time and Boolean types are automatically identified and serialized to your spreadsheet.

5. Automatic and fixed column widths: Axlsx will automatically determine the appropriate width for your columns based on the content in the worksheet, or use any value you specify for the really funky stuff.

6. Support for automatically formatted 1904 and 1900 epochs configurable in the workbook.

7. Add jpg, gif and png images to worksheets with hyperlinks

8. Reference cells in your worksheet with "A1" and "A1:D4" style references or from the workbook using "Sheet1!A3:B4" style references

9. Cell level style overrides for default and customized style objects

10. Support for formulas, merging, row and column outlining as well as
cell level input data validation.

12. Auto filtering tables with worksheet.auto_filter as well as support for Tables

13. Export using shared strings or inline strings so we can inter-op with iWork Numbers (sans charts for now).

14. Output to file or StringIO

15. Support for page margins and print options

16. Support for password and non password based sheet protection.

17. First stage interoperability support for GoogleDocs, LibreOffice,
and Numbers

18. Support for defined names, which gives you repeated header rows for printing.

19. Data labels for charts as well as series color customization.

20. Support for sheet headers and footers

21. Pivot Tables

22. Page Breaks

Installing
----------

To install Axlsx, use the following command:

    $ gem install axlsx

#Examples
------

The example listing is getting overly large to maintain here.
If you are using Yard, you will be able to see the examples in line below.

Here's a teaser that kicks about 2% of what the gem can do.

```ruby
Axlsx::Package.new do |p|
  p.workbook.add_worksheet(:name => "Pie Chart") do |sheet|
    sheet.add_row ["Simple Pie Chart"]
    %w(first second third).each { |label| sheet.add_row [label, rand(24)+1] }
    sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,5], :end_at => [10, 20], :title => "example 3: Pie Chart") do |chart|
      chart.add_series :data => sheet["B2:B4"], :labels => sheet["A2:A4"],  :colors => ['FF0000', '00FF00', '0000FF']
    end
  end
  p.serialize('simple.xlsx')
end
```

Please see the [examples](https://github.com/randym/axlsx/tree/master/examples/example.rb) for more.

{include:file:examples/example.rb}

There is much, much more you can do with this gem. If you get stuck, grab me on IRC or submit an issue to GitHub. Chances are that it has already been implemented. If it hasn't - let's take a look at adding it in.

#Documentation
--------------

This gem is 100% documented with YARD, an exceptional documentation library. To see documentation for this, and all the gems installed on your system use:

     gem install yard

     yard server -g

#Specs
------

This gem has 100% test coverage using test/unit. To execute tests for this gem, simply run rake in the gem directory.

#Change log
---------

- **September.12.13**:2.0.1
  - Unpinned rubyzip version
- **September.12.13**:2.0.0
  - DROPPED support for ruby 1.8.7
  - Altered readme to link to contributors
  - Lots of improvements to make charts and relations more stable.
  - Patched color param mutation.
  - Data sourced for pivot tables can now come from other sheets.
  - Altered image file extension comparisons to be case insensitive.
  - Added control character sanitization to shared strings.
  - Added page breaks. see examples/example.rb for an example.
  - Bugfix: single to dual cell anchors for images now swap properly so you can set the end_at position during instantiation, in a block or directly on the image.
  - Improved how we convert date/time to include the UTC offset when provided.
  - Pinned rubyzip to 0.9.9 for those who are not ready to go up. Please note that release 2.0.1 and on will be using the 1.n.n series of rubyzip
  - Bugfix: transposition of cells for Worksheet#cols now supports
    incongruent column counts.counts
  - Added space preservation for cell text. This will allow whitespace
    in cell text both when using shared strings and when serializing
    directly to the cell.
- **April.24.13**:1.3.6
  - Fixed LibreOffice/OpenOffice issue to properly apply colors to lines
    in charts.
  - Added support for specifying between/notBetween formula in an array.
    *thanks* straydogstudio!
  - Added standard line chart support. *thanks* scambra
  - Fixed straydogstudio's link in the README. *thanks* nogara!
- **February.4.13**:1.3.5
  - converted vary_colors for chart data to instance variable with appropriate defulats for the various charts.
  - Added trust_input method on Axlsx to instruct the serializer to skip HTML escaping. This will give you a tremendous performance boost,
    Please be sure that you will never have <, >, etc in your content or the XML will be invalid.
  - Rewrote cell serialization to improve performance
  - Added iso_8601 type to support text based date and time management.
  - Bug fix for relationahip management in drawings when you add images
    and charts to the same worksheet drawing.
  - Added outline_level_rows and outline_level_columns to worksheet to simplify setting up outlining in the worksheet.
  - Added support for pivot tables
  - Added support for descrete border edge styles
  - Improved validation of sheet names
  - Added support for formula value caching so that iOS and OSX preview can show the proper values. See Cell.add_row and the formula_values option.

Please see the {file:CHANGELOG.md} document for past release information.

# Known interoperability issues.
As axslx implements the Office Open XML (ECMA-376 spec) much of the
functionality is interoperable with other spreadsheet software. Below is
a listing of some known issues.

1. Libre Office
   -  You must specify colors for your series. see examples/chart_colors.rb
for an example.
   - You must use data in your sheet for charts. You cannot use hard coded
values.
   -  Chart axis and gridlines do not render. I have a feeling this is
related to themes, which axlsx does not implement at this time.

2. Google Docs
   - Images are known to not work with google docs
   - border colors do not work

3. Numbers
   - you must set 'use_shared_strings' to true. This is most
     conveniently done just before rendering by calling Package.use_shared_strings = true prior to serialization.

```ruby
p = Axlsx::Package.new
p.workbook.add_worksheet(:name => "Basic Worksheet") do |sheet|
  sheet.add_row ["First Column", "Second", "Third"]
  sheet.add_row [1, 2, 3]
end
p.use_shared_strings = true
p.serialize('simple.xlsx')
```

   - charts do not render


#Thanks!

Open source software is a community effort. None of this could have been
done without the help of these awesome folks.

[contributors](https://github.com/randym/axlsx/graphs/contributors)

#Copyright and License
----------

Axlsx &copy; 2011-2013 by [Randy Morgan](mailto:digial.ipseity@gmail.com).

Axlsx is licensed under the MIT license. Please see the LICENSE document for more information.
