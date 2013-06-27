$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
p = Axlsx::Package.new
p.workbook do |wb|
  wb.styles do |s|
    no_underline = s.add_style :sz => 10, :b => true, :u => false, :alignment => { :horizontal=> :right }
    wb.add_worksheet(:name => 'wunderlinen') do |sheet|
      sheet.add_row %w{a b c really?}, :style => no_underline
    end
  end
end
p.serialize 'no_underline.xlsx'

