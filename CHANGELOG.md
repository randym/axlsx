CHANGELOG
---------
- **November.25.12**:1.3.4
  - Support for headers and footers for worksheets
  - bug fix: Properly escape hyperlink urls
  - Improvements in color_scale generation for conditional formatting
  - Improvements in autowidth calculation.
- **November.8.12**:1.3.3
  - Patched cell run styles for u and validation for family

- **November.5.12**:1.3.2
  - MASSIVE REFACTORING
  - Patch for apostrophes in worksheet names
  - added sheet_by_name for workbook so you can now find your worksheets
    by name
  - added insert_worksheet so you can now add a worksheet to an
    arbitrary position in the worksheets list.
  - reduced memory consumption for package parts post serialization
- **September.30.12**: 1.3.1
  - Improved control character handling
  - Added stored auto filter values and date grouping items
  - Improved support for autowidth when custom styles are applied
  - Added support for table style info that lets you take advantage of
    all the predefined table styles.
  - Improved style management for fonts so they merge undefined values
    from the initial master.
- **September.8.12**: 1.2.3
  - enhance exponential float/bigdecimal values rendering as strings intead
    of 'numbers' in excel.
  - added support for :none option on chart axis labels
  - added support for paper_size option on worksheet.page_setup
- **August.27.12**: 1.2.2
   - minor patch for auto-filters
   - minor documentation improvements.
- **August.12.12**: 1.2.1
   - Add support for hyperlinks in worksheets
   - Fix example that was using old style cell merging with concact.
   - Fix bug that occurs when calculating the font_size for cells that use a user specified style which does not define sz
- **August.5.12**: 1.2.0
   - rebuilt worksheet serialization to clean up the code base a bit.
   - added data labels for charts
   - added table header printing for each sheet via defined_name. This
     means that when you print your worksheet, the header rows show for every page
- **July.??.12**: 1.1.9 release
   - lots of code clean up for worksheet
   - added in data labels for pie charts, line charts and bar charts.
   - bugfix chard with data in a sheet that has a space in the name are
     now auto updating formula based values
- **July.14.12**: 1.1.8 release
   - added html entity encoding for sheet names. This allows you to use
     characters like '<' and '&' in your sheet names.
   - new - first round google docs interoperability
   - added filter to strip out control characters from cell data.
   - added in interop requirements so that charts are properly exported
     to PDF from Libra Office
   - various readability improvements and work standardizing attribute
     names to snake_case. Aliases are provided for backward compatiblity 
- **June.11.12**: 1.1.7 release
   - fix chart rendering issue when label offset is specified as a
     percentage in serialization and ensure that formula are not stored
in value caches
   - fix bug that causes repair warnings when using a text only title reference.
   - Add title property to axis so you can lable the x/y/series axis for
     charts.
   - Add sheet views with panes
- **May.30.12**: 1.1.6 release
   - data protection with passwords for sheets
   - cell level input validators
   - added support for two cell anchors for images
   - test coverage now back up to 100%
   - bugfix for merge cell sorting algorithm
   - added fit_to method for page_setup to simplify managing witdh/height
   - added ph (phonetics) and s (style) attributes for row.
   - resolved all warnings generating from this gem.
   - improved comment relationship management for multiple comments
- **May.13.12**: 1.1.5 release
   - MOAR print options! You can now specify paper size, orientation,
     fit to width, page margings and gridlines for printing.
   - Support for adding comments to your worksheets
   - bugfix for applying style to empty cells
   - bugfix for parsing formula with multiple '='

- **May.3.12:**: 1.1.4 release
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

- ** March.5.12**: 1.0.18 release
   https://github.com/randym/axlsx/compare/1.0.17...1.0.18
   - bugfix custom borders are not properly applied when using styles.add_style
   - interop worksheet names must be 31 characters or less or some versions of office complain about repairs
   - added type support for :boolean and :date types cell values
   - added support for fixed column widths
   - added support for page_margins
   - added << alias for add_row
   - removed presetting of date1904 based on authoring platform. Now defaults to use 1900 epoch (date1904 = false)

- ** February.14.12**: 1.0.17 release
   https://github.com/randym/axlsx/compare/1.0.16...1.0.17
   - Added in support for serializing to StringIO
   - Added in support for using shared strings table. This makes most of the features in axlsx interoperable with iWorks Numbers
   - Added in support for fixed column_widths
   - Removed unneeded dependencies on active-support and i18n

- ** February.2.12**: 1.0.16 release
   https://github.com/randym/axlsx/compare/1.0.15...1.0.16
   - Bug fix for schema file locations when validating in rails
   - Added hyperlink to images
   - date1904 now automatically set in BSD and mac environments
   - removed whitespace/indentation from xml outputs
   - col_style now skips rows that do not contain cells at the column index

- **January.6.12**: 1.0.15 release
   https://github.com/randym/axlsx/compare/1.0.14...1.0.15
   - Bug fix add_style specified number formats must be explicity applied for libraOffice
   - performance improvements from ochko when creating cells with options.
   - Bug fix setting types=>[:n] when adding a row incorrectly determines the cell type to be string as the value is null during creation.
   - Release in preparation for password protection merge

- **December.14.11**: 1.0.14 release
   - Added support for merging cells
   - Added support for auto filters
   - Improved auto width calculations
   - Improved charts
   - Updated examples to output to a single workbook with multiple sheets
   - Added access to app and core package objects so you can set the creator and other properties of the package
   - The beginning of password protected xlsx files - roadmapped for January release.

- **December.8.11**: 1.0.13 release
   -  Fixing .gemspec errors that caused gem to miss the lib directory. Sorry about that.

- **December.7.11**: 1.0.12 release
    DO NOT USE THIS VERSION = THE GEM IS BROKEN
  - changed dependency from 'zip' gem to 'rubyzip' and added conditional code to force binary encoding to resolve issue with excel 2011
  - Patched bug in app.xml that would ignore user specified properties.
- **December.5.11**: 1.0.11 release
  - Added [] methods to worksheet and workbook to provide name based access to cells.
  - Added support for functions as cell values
  - Updated color creation so that two character shorthand values can be used like 'FF' for 'FFFFFFFF' or 'D8' for 'FFD8D8D8'
  - Examples for all this fun stuff added to the readme
  - Clean up and support for 1.9.2 and travis integration
  - Added support for string based cell referencing to chart start_at and end_at. That means you can now use :start_at=>"A1" when using worksheet.add_chart, or chart.start_at ="A1" in addition to passing a cell or the x, y coordinates.

- **October.30.11**: 1.0.10 release
  - Updating gemspec to lower gem version requirements.
  - Added row.style assignation for updating the cell style for an entire row
  - Added col_style method to worksheet upate a the style for a column of cells
  - Added cols for an easy reference to columns in a worksheet.
  - prep for pre release of acts_as_xlsx gem
  - added in travis.ci configuration and build status
  - fixed out of range bug in time calculations for 32bit time.
  - added i18n for active support

- **October.26.11**: 1.0.9 release
  - Updated to support ruby 1.9.3
  - Updated to eliminate all warnings originating in this gem

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

##October.22.11: 1.0.6 release
  - Bumping version to include docs. Bug in gemspec pointing to incorrect directory.

##October.22.11: 1.05
  - Added support for line charts
  - Updated examples and readme
  - Updated series title to be a real title ** NOTE ** If you are accessing titles directly you will need to update text assignation.
         chart.series.first.title = 'Your Title'
         chart.series.first.title.text = 'Your Title'
    With this change you can assign a cell for the series title
         chart.series.title = sheet.rows.first.cells.first
    If you are using the recommended
         chart.add_series :data=>[], :labels=>[], :title
    You do not have to change anything.
  - BugFix: shape attribute for bar chart is now properly serialized
  - BugFix: date1904 property now properly set for charts
  - Added style property to charts
  - Removed serialization write test as it most commonly fails when run from the gem's intalled directory

##October.21.11: 1.0.4
  - altered package to accept a filename string for serialization instead of a File object.
  - Updated specs to conform
  - More examples for readme


##October.21.11: 1.0.3 release
  - Updated documentation

##October.20.11: 0.1.0 release

