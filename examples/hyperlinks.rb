#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

#```ruby
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook
wb.add_worksheet(:name => 'hyperlinks') do |sheet|
  # external references
  sheet.add_row ['axlsx']
  sheet.add_hyperlink :location => 'https://github.com/randym/axlsx', :ref => sheet.rows.first.cells.first
  # internal references
  sheet.add_row ['next sheet']
  sheet.add_hyperlink :location => "'Next Sheet'!A1", :target => :sheet, :ref => 'A2'
end

wb.add_worksheet(:name => 'Next Sheet') do |sheet|
  sheet.add_row ['hello!']
end

p.serialize 'hyperlinks.xlsx'
