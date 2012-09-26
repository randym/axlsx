#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-

$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
Axlsx::Package.new do |p|
  p.workbook.add_worksheet(:name => "Table") do |sheet|
    sheet.add_row ["Build Matrix"]
    sheet.add_row ["Build", "Duration", "Finished", "Rvm"]
    sheet.add_row ["19.1", "1 min 32 sec", "about 10 hours ago", "1.8.7"]
    sheet.add_row ["19.2", "1 min 28 sec", "about 10 hours ago", "1.9.2"]
    sheet.add_row ["19.3", "1 min 35 sec", "about 10 hours ago", "1.9.3"]
    sheet.auto_filter = 'A2:D5' 
    sheet.auto_filter.add_column 3, :filters, :filter_items => ['1.9.2']
  end
end.serialize('auto_filter.xlsx')
