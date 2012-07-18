Axlsx – The end of CSV as an excuse for client reporting. 

One of the things that keeps popping up for ruby and ruby on rails application developers is client reporting.
We have tools for creating beautiful screen based reports that make the end user say “WOW”, we can even - 
with the help of a great designer - produce a well formatted PDF. But when it comes to data, we are still
using CSV, a technology from the 1960's

That has changed. There is a new gem, released in November of last year and still under active development
that brings more to ruby and ruby on rails than was ever possible before for client reporting.

The gem is axlsx (http://rubygems.org/gems/axlsx) 

Let's take a look at some of the things that it can do for you.

1. The Basics
The basics of generating and serializing xlsx data are dead simple. 

1. Create a package
2. Setup your styles
3. add a worksheet to the workbook
3. add your data, charts, images, conditional formatting, data
   validations, page setup, print options, and password locking,
comments, cell merges and panes. 
4. serialize

The code below illustrates a trivial example a few of these items

'''ruby
require 'axlsx'

package = Axlsx::Package.new
styles = package.workbook.styles

header = styles.add_style :bg_color => '00', :fg_color => 'FF', :sz => 16, :alignment => { :horizontal => :center } 

quarter_label = styles.add_style :bg_color => 'FFDFDEDF', :alignment => { :indent => 1 }, :sz => 14

money = styles.add_style :num_fmt => 5, :zd => 14

package.workbook.add_worksheet do |worksheet|
  worksheet.add_row ['Revenue by Quarter'], :style => header
  worksheet.merge_cells 'A1:I1'
  worksheet.add_row
  data_row_style = [nil, quarter_label, money]

  worksheet.add_row [nil, 'Q1', 35221124], :style => data_row_style
  worksheet.add_row [nil, 'Q2', 56742113], :style => data_row_style
  worksheet.add_row [nil, 'Q3', 71165443], :style => data_row_style
  worksheet.add_row [nil, 'Q4', 98761111], :style => data_row_style

  worksheet.add_chart(Axlsx::Bar3DChart, :bar_dir => :col) do |chart|
    chart.start_at 4, 2
    chart.end_at 9, 15
    chart.title = worksheet['A1']
  chart.add_series :data => worksheet['C3:C6'], :labels => worksheet['B3:B6']
  end
end
package.serialize 'the_better_basics.xlsx'

If you are using rails, there is a sister gem acts_as_xlsx (http://rubygems.org/gems/axlsx)
That is going to make you smile as well.

Have a look at these two blog posts for a quick write up of how easy it
is to use.

http://axlsx.blogspot.jp/2011/12/using-actsasxlsx-to-generate-excel-data.html
http://axlsx.blogspot.jp/2011/12/axlsx-making-excel-reports-with-ruby-on.html



 

 

