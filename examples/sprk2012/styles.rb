require 'axlsx'
package = Axlsx::Package.new
workbook = package.workbook

workbook.styles do |s|

  black_cell = s.add_style :bg_color => "00", :fg_color => "FF", :sz => 14, :alignment => { :horizontal=> :center }
  blue_cell =  s.add_style  :bg_color => "0000FF", :fg_color => "FF", :sz => 20, :alignment => { :horizontal=> :center }

  wb.add_worksheet(:name => "Styles") do |sheet|

    # Applies the black_cell style to the first and third cell, and the blue_cell style to the second.
    sheet.add_row ["Text Autowidth", "Second", "Third"], :style => [black_cell, blue_cell, black_cell]

    # Applies the thin border to all three cells
    sheet.add_row [1, 2, 3], :style => Axlsx::STYLE_THIN_BORDER
  end
end

