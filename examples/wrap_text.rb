$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
p = Axlsx::Package.new
p.workbook do |wb|
  wb.styles do |s|
    wrap_text = s.add_style :fg_color=> "FFFFFF",
                            :b => true,
                            :bg_color => "004586",
                            :sz => 12,
                            :border => { :style => :thin, :color => "00" },
                            :alignment => { :horizontal => :center,
                                            :vertical => :center ,
                                            :wrap_text => true}
    wb.add_worksheet(:name => 'wrap text') do |sheet|
      sheet.add_row ['Torp, White and Cronin'], :style => wrap_text
      # Forcing the column to be a bit narrow so we can see if the text wrap.
      sheet.column_info.first.width = 5
    end
  end
end
p.serialize 'wrap_text.xlsx'
