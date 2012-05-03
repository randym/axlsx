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

**Latest Version**: 1.1.4

**Ruby Version**: 1.8.7, 1.9.2, 1.9.3

**JRuby Version**: 1.6.7 1.8 and 1.9 modes

**Rubinius Version**: rubinius 2.0.0dev * lower versions may run, this gem always tests against head.

**Release Date**: May 3rd 2012

If you are working in rails, or with active record see:
http://github.com/randym/acts_as_xlsx

There are guides for using axlsx and acts_as_xlsx here:
[http://axlsx.blogspot.com](http://axlsx.blogspot.com)

Synopsis
--------

Axlsx is an Office Open XML Spreadsheet generator for the Ruby programming language.
With Axlsx you can create excel worksheets with charts, images (with links), automated and fixed column widths, customized styles, functions, merged cells, auto filters, file and stream serialization  as well as full schema validation. Axlsx excels at helping you generate beautiful Office Open XML Spreadsheet documents without having to understand the entire ECMA specification.

![Screen 1](https://github.com/randym/axlsx/raw/master/examples/sample.png)



Feature List
------------

**1. Author xlsx documents: Axlsx is made to let you easily and quickly generate professional xlsx based reports that can be validated before serialization.

**2. Generate 3D Pie, Line, Scatter and Bar Charts: With Axlsx chart generation and management is as easy as a few lines of code. You can build charts based off data in your worksheet or generate charts without any data in your sheet at all. Customize gridlines, label rotation and series colors as well.

**3. Custom Styles: With guaranteed document validity, you can style borders, alignment, fills, fonts, and number formats in a single line of code. Those styles can be applied to an entire row, or a single cell anywhere in your workbook.

**4. Automatic type support: Axlsx will automatically determine the type of data you are generating. In this release Float, Integer, String, Date, Time and Boolean types are automatically identified and serialized to your spreadsheet.

**5. Automatic and fixed column widths: Axlsx will automatically determine the appropriate width for your columns based on the content in the worksheet, or use any value you specify for the really funky stuff.

**6. Support for automatically formatted 1904 and 1900 epochs configurable in the workbook.

**7. Add jpg, gif and png images to worksheets with hyperlinks

**8. Reference cells in your worksheet with "A1" and "A1:D4" style references or from the workbook using "Sheet1!A3:B4" style references

**9. Cell level style overrides for default and customized style objects

**10. Support for formulas

**11. Support for cell merging as well as column and row outline

**12. Auto filtering tables with worksheet.auto_filter as well as support for Tables

**13. Export using shared strings or inline strings so we can inter-op with iWork Numbers (sans charts for now).

**14. Output to file or StringIO

**15. Support for page margins and print options

Installing
----------

To install Axlsx, use the following command:

    $ gem install axlsx

#Examples
------

The example listing is getting overly large to maintain here.
If you are using Yard, you will be able to see the examples inline below.
If not, please refer to the Please see the {file:examples/example.rb} file.

{include:file:examples/example.rb}

There is much, much more you can do with this gem. If you get stuck, grab me on IRC or submit an issue to Github. Chances are that it has already been implemented. If it hasn't - let's take a look at adding it in.

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
- ** May.3.12:**: 1.1.4 release
   - MOAR examples
   - added outline level for rows and columns
   - rebuild of numeric and axis data sources for charts
   - added delete to axis
   - added tick and label mark skipping for cat axis in charts
   - bugfix for table headers method
   - sane(er) defaults for chart positioning
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
