$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook
page_header = wb.styles.add_style :bg_color => "DD", :sz => 16, :b => true, :alignment => {:horizontal => :center}
table_title = wb.styles.add_style :b => true, :alignment => { :horizontal => :center }
col_header  = wb.styles.add_style :bg_color => "FF95A39D", :b => true, :alignment => {:indent => 1}
label       = wb.styles.add_style :alignment => { :indent => 1 }
money       = wb.styles.add_style :num_fmt => 5
# The first sheet in the workbook is a summary of data from the second sheet.
wb.add_worksheet do |sheet|
  sheet.add_row
  sheet.add_row [nil, "College Budget"], :style => [nil, page_header]
  sheet.merged_cells << "B2:F2"
  sheet.add_row
  sheet.add_row [nil, "What's coming in this month.", nil, nil, "How am I doing"], :style => table_title
  sheet.merged_cells << "B4:C4"
  sheet.merged_cells << "E4:F4"
  sheet.add_row [nil, "Item", "Amount", nil, "Item", "Amount"], :style => col_header
  sheet.add_row [nil, "Estimated monthly net income", 500, nil, "Monthly income", "=C9"], :style => [nil, label, money, nil, label, money]
  sheet.add_row [nil, "Financial Awards", 100, nil, "Monthly expenses", "=C27"], :style =>  [nil, label, money, nil, label, money]
  sheet.add_row [nil, "Allowance from mom & dad", 20000, nil, "Semester expenses", "=F19"], :style =>  [nil, label, money, nil, label, money]
  sheet.add_row [nil, "Total", "=SUM(C6:C8)", nil, "Difference", "=F6 - SUM(F7:F8)"]
  sheet.add_row
  sheet.add_row [nil, "What's going out this month.", nil, nil, "Semester Costs"], :style => table_title
  sheet.merged_cells << "B11:C11"
  sheet.merged_cells << "E11:F11"
  sheet.add_row [nil, "Item", "Amount", nil, "Item", "Amount"], :style => col_header
  sheet.add_row [nil, "Rent", 650, nil, "Tuition", 200], :style =>  [nil, label, money, nil, label, money]
  sheet.add_row [nil, "Utilities", 120, nil, "Lab fees", 50], :style =>  [nil, label, money, nil, label, money]
  sheet.add_row [nil, "Cell phone", 100, nil, "Other fees", 10], :style =>  [nil, label, money, nil, label, money]
  sheet.add_row [nil, "Groceries", 75, nil, "Books", 150], :style =>  [nil, label, money, nil, label, money]
  sheet.add_row [nil, "Auto expenses", 0, nil, "Deposits", 0], :style =>  [nil, label, money, nil, label, money]
  sheet.add_row [nil, "Student loans", 0, nil, "Transportation", 30], :style =>  [nil, label, money, nil, label, money]
  sheet.add_row [nil, "Other loans", 350, nil, "Total", "=SUM(F13:F18)"]
  sheet.add_row [nil, "Credit cards", 450]
  sheet.add_row [nil, "Insurance", 0]
  sheet.add_row [nil, "Laundry", 10]
  sheet.add_row [nil, "Haircuts", 0]
  sheet.add_row [nil, "Medical expenses", 0]
  sheet.add_row [nil, "Entertainment", 500]
  sheet.add_row [nil, "Miscellaneous", 0]
  sheet.add_row [nil, "Total", "=SUM(C13:C26)"]
  sheet.add_chart(Axlsx::Pie3DChart) do |chart|
    chart.title = sheet["B11"]
    chart.add_series :data => sheet["C13:C26"], :labels => sheet["B13:B26"]
    chart.start_at 7, 2
    chart.end_at 12, 15
  end
  sheet.add_chart(Axlsx::Bar3DChart, :barDir => :col) do |chart|
    chart.title = sheet["E11"]
    chart.add_series :labels => sheet["E13:E18"], :data => sheet["F13:F18"]
    chart.start_at 7, 16
    chart.end_at 12, 31
  end
  sheet.column_widths 2, 23, nil, 2, 23, nil, 2
end

p.serialize 'real_example.xlsx'
