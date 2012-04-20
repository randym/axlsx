#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'

p = Axlsx::Package.new
book = p.workbook
ws = book.add_worksheet

# define your regular styles
percent = book.styles.add_style(:format_code => "0.00%", :border => Axlsx::STYLE_THIN_BORDER)
money = book.styles.add_style(:format_code => '0,000', :border => Axlsx::STYLE_THIN_BORDER)

# define the style for conditional formatting
profitable = book.styles.add_style( :fg_color=>"FF428751",
                                    :type => :dxf)

# Generate 20 rows of data
ws.add_row ["Previous Year Quarterly Profits (JPY)"]
ws.add_row ["Quarter", "Profit", "% of Total"]
offset = 3
rows = 20
offset.upto(rows + offset) do |i|
  ws.add_row ["Q#{i}", 10000*((rows/2-i) * (rows/2-i)), "=100*B#{i}/SUM(B3:B#{rows+offset})"], :style=>[nil, money, percent]
end

# Apply conditional formatting to range B4:B100 in the worksheet
ws.add_conditional_formatting("B4:B100", { :type => :cellIs, :operator => :greaterThan, :formula => "100000", :dxfId => profitable, :priority => 1 })

f = File.open('example_differential_styling.xlsx', 'w')
p.serialize(f)
