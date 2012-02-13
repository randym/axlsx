CHANGELOG
---------
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
