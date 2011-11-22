CHANGELOG
---------

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
