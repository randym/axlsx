require 'axlsx' 

############################### 
# Function to output results data row to summary spreadsheet 
def outputRow (sid, type) 

  $sumSheet.add_row [ sid, type, "1", "2", "3", "4", "5" ], :style => $black_cell 

  if sid.odd?
    link = "A#{$curRow}" 
    puts "outputRow: sid: #{sid}, link: #{link}" 
    # Setting the style for the link will apply the xf that we created in the Main Program block
    $sumSheet[link].style = $blue_link
    $sumSheet.add_hyperlink :location => "'Log'!A#{$curRow}", :target => :sheet, :ref => link 
  end
  $curRow += 1 
end 

############################## 
#   Main Program 

$package = Axlsx::Package.new 
$workbook = $package.workbook 
##  We want to create our sytles outside of the outputRow method
#   Each style only needs to be declared once in the workbook.
$workbook.styles do |s|
  $black_cell = s.add_style :sz => 10, :alignment => { :horizontal=> :center } 
  $blue_link = s.add_style :fg_color => '0000FF'
end 


# Create summary sheet 
$sumSheet = $workbook.add_worksheet(:name => 'Summary') 
$sumSheet.add_row ["Test Results"], :sz => 16 
$sumSheet.add_row 
$sumSheet.add_row 
$sumSheet.add_row ["Note: Blue cells below are links to the Log sheet"], :sz => 10 
$sumSheet.add_row 
$workbook.styles do |s| 
  black_cell = s.add_style :sz => 14, :alignment => { :horizontal=> :center } 
  $sumSheet.add_row ["ID","Type","Match","Mismatch","Diffs","Errors","Result"], :style => black_cell 
end 
$sumSheet.column_widths 10, 10, 10, 11, 10, 10, 10 

# Starting data row in summary spreadsheet (after header info) 
$curRow = 7 

# Create Log Sheet 
$logSheet = $workbook.add_worksheet(:name => 'Log') 
$logSheet.column_widths 10 
$logSheet.add_row ['Log Detail'], :sz => 16 
$logSheet.add_row 

# Add rows to summary sheet 
for i in 1..10 do 
  outputRow(i, 'test') 
end 

$package.serialize 'where_is_my_color.xlsx'
