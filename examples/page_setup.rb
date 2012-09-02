$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
xls = Axlsx::Package.new
wb = xls.workbook
wb.add_worksheet do |ws|
  # Excel does not currently follow the specification and will ignore paper_width and paper_height so if you need
  # to for a paper size, be sure to set :paper_size
  ws.page_setup.set :paper_width => "210mm", :paper_size => 10,  :paper_height => "297mm", :orientation => :landscape
  ws.add_row %w(AXLSX is cool)
end
xls.serialize "page_setup.xlsx"
