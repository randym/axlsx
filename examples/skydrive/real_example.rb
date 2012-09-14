$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../../lib"

require 'axlsx'

Axlsx::Package.new do |p|
  p.workbook do |wb|
    styles = wb.styles
    header     = styles.add_style :bg_color => "DD", :sz => 16, :b => true, :alignment => {:horizontal => :center}
    tbl_header = styles.add_style :b => true, :alignment => { :horizontal => :center }
    ind_header = styles.add_style :bg_color => "FFDFDEDF", :b => true, :alignment => {:indent => 1}
    col_header = styles.add_style :bg_color => "FFDFDEDF", :b => true, :alignment => { :horizontal => :center }
    label      = styles.add_style :alignment => { :indent => 1 }
    money      = styles.add_style :num_fmt => 5
    t_label    = styles.add_style :b => true, :bg_color => "FFDFDEDF"
    t_money    = styles.add_style :b => true, :num_fmt => 5, :bg_color => "FFDFDEDF"

    wb.add_worksheet do |sheet|
      sheet.add_row
      sheet.add_row [nil, "College Budget"], :style => [nil, header]
      sheet.add_row
      sheet.add_row [nil, "What's coming in this month.", nil, nil, "How am I doing"], :style => tbl_header
      sheet.add_row [nil, "Item", "Amount", nil, "Item", "Amount"], :style => [nil, ind_header, col_header, nil, ind_header, col_header]
      sheet.add_row [nil, "Estimated monthly net income", 500, nil, "Monthly income", "=C9"], :style => [nil, label, money, nil, label, money]
      sheet.add_row [nil, "Financial aid", 100, nil, "Monthly expenses", "=C27"], :style =>  [nil, label, money, nil, label, money]
      sheet.add_row [nil, "Allowance from mom & dad", 20000, nil, "Semester expenses", "=F19"], :style =>  [nil, label, money, nil, label, money]
      sheet.add_row [nil, "Total", "=SUM(C6:C8)", nil, "Difference", "=F6 - SUM(F7:F8)"], :style => [nil, t_label, t_money, nil, t_label, t_money]
      sheet.add_row
      sheet.add_row [nil, "What's going out this month.", nil, nil, "Semester Costs"], :style => tbl_header
      sheet.add_row [nil, "Item", "Amount", nil, "Item", "Amount"], :style => [nil, ind_header, col_header, nil, ind_header, col_header]
      sheet.add_row [nil, "Rent", 650, nil, "Tuition", 200], :style =>  [nil, label, money, nil, label, money]
      sheet.add_row [nil, "Utilities", 120, nil, "Lab fees", 50], :style =>  [nil, label, money, nil, label, money]
      sheet.add_row [nil, "Cell phone", 100, nil, "Other fees", 10], :style =>  [nil, label, money, nil, label, money]
      sheet.add_row [nil, "Groceries", 75, nil, "Books", 150], :style =>  [nil, label, money, nil, label, money]
      sheet.add_row [nil, "Auto expenses", 0, nil, "Deposits", 0], :style =>  [nil, label, money, nil, label, money]
      sheet.add_row [nil, "Student loans", 0, nil, "Transportation", 30], :style =>  [nil, label, money, nil, label, money]
      sheet.add_row [nil, "Other loans", 350, nil, "Total", "=SUM(F13:F18)"], :style => [nil, label, money, nil, t_label, t_money]
      sheet.add_row [nil, "Credit cards", 450], :style => [nil, label, money]
      sheet.add_row [nil, "Insurance", 0], :style => [nil, label, money]
      sheet.add_row [nil, "Laundry", 10], :style => [nil, label, money]
      sheet.add_row [nil, "Haircuts", 0], :style => [nil, label, money]
      sheet.add_row [nil, "Medical expenses", 0], :style => [nil, label, money]
      sheet.add_row [nil, "Entertainment", 500], :style => [nil, label, money]
      sheet.add_row [nil, "Miscellaneous", 0], :style => [nil, label, money]
      sheet.add_row [nil, "Total", "=SUM(C13:C26)"], :style => [nil, t_label, t_money]
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
      %w(B4:C4 E4:F4 B11:C11 E11:F11 B2:F2).each { |range| sheet.merge_cells(range) }
      sheet.column_widths 2, nil, nil, 2, nil, nil, 2
    end
  end
  p.use_shared_strings = true
  p.serialize 'axlsx.xlsx'
end
