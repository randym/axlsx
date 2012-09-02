$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
xls = Axlsx::Package.new
wb = xls.workbook
wb.add_worksheet do |ws|
  ws.page_setup.set :paper_width => "210mm", :paper_size => 10,  :paper_height => "297mm", :orientation => :landscape
  ws.add_row %w(AXLSX is cool)
end
xls.serialize "page_setup.xlsx"
