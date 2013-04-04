#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'

p = Axlsx::Package.new
book = p.workbook

# define your regular styles
percent = book.styles.add_style(:format_code => "0.00%", :border => Axlsx::STYLE_THIN_BORDER)
money = book.styles.add_style(:format_code => '0,000', :border => Axlsx::STYLE_THIN_BORDER)

# define the style for conditional formatting
profitable = book.styles.add_style( :fg_color => "FF428751", :type => :dxf )
unprofitable = wb.styles.add_style( :fg_color => "FF0000", :type => :dxf )

book.add_worksheet(:name => "Cell Is") do |ws|

  # Generate 20 rows of data
  ws.add_row ["Previous Year Quarterly Profits (JPY)"]
  ws.add_row ["Quarter", "Profit", "% of Total"]
  offset = 3
  rows = 20
  offset.upto(rows + offset) do |i|
    ws.add_row ["Q#{i}", 10000*((rows/2-i) * (rows/2-i)), "=100*B#{i}/SUM(B3:B#{rows+offset})"], :style=>[nil, money, percent]
  end

# Apply conditional formatting to range B3:B100 in the worksheet
  ws.add_conditional_formatting("B3:B100", { :type => :cellIs, :operator => :greaterThan, :formula => "100000", :dxfId => profitable, :priority => 1 })
# Apply conditional using the between operator; NOTE: supply an array to :formula for between/notBetween
  sheet.add_conditional_formatting("C3:C100", { :type => :cellIs, :operator => :between, :formula => ["0.00%","100.00%"], :dxfId => unprofitable, :priority => 1 })
end

book.add_worksheet(:name => "Color Scale") do |ws|
  ws.add_row ["Previous Year Quarterly Profits (JPY)"]
  ws.add_row ["Quarter", "Profit", "% of Total"]
  offset = 3
  rows = 20
  offset.upto(rows + offset) do |i|
    ws.add_row ["Q#{i}", 10000*((rows/2-i) * (rows/2-i)), "=100*B#{i}/SUM(B3:B#{rows+offset})"], :style=>[nil, money, percent]
  end
# Apply conditional formatting to range B3:B100 in the worksheet
  color_scale = Axlsx::ColorScale.new
  ws.add_conditional_formatting("B3:B100", { :type => :colorScale, :operator => :greaterThan, :formula => "100000", :dxfId => profitable, :priority => 1, :color_scale => color_scale })
end


book.add_worksheet(:name => "Data Bar") do |ws|
  ws.add_row ["Previous Year Quarterly Profits (JPY)"]
  ws.add_row ["Quarter", "Profit", "% of Total"]
  offset = 3
  rows = 20
  offset.upto(rows + offset) do |i|
    ws.add_row ["Q#{i}", 10000*((rows/2-i) * (rows/2-i)), "=100*B#{i}/SUM(B3:B#{rows+offset})"], :style=>[nil, money, percent]
  end
# Apply conditional formatting to range B3:B100 in the worksheet
  data_bar = Axlsx::DataBar.new
  ws.add_conditional_formatting("B3:B100", { :type => :dataBar, :dxfId => profitable, :priority => 1, :data_bar => data_bar })
end

book.add_worksheet(:name => "Icon Set") do |ws|
  ws.add_row ["Previous Year Quarterly Profits (JPY)"]
  ws.add_row ["Quarter", "Profit", "% of Total"]
  offset = 3
  rows = 20
  offset.upto(rows + offset) do |i|
    ws.add_row ["Q#{i}", 10000*((rows/2-i) * (rows/2-i)), "=100*B#{i}/SUM(B3:B#{rows+offset})"], :style=>[nil, money, percent]
  end
# Apply conditional formatting to range B3:B100 in the worksheet
  icon_set = Axlsx::IconSet.new
  ws.add_conditional_formatting("B3:B100", { :type => :iconSet, :dxfId => profitable, :priority => 1, :icon_set => icon_set })
end

p.serialize('example_conditional_formatting.xlsx')
