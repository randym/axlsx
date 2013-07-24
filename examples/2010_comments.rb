#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

#```ruby
require 'axlsx'
p = Axlsx::Package.new
p.workbook.add_worksheet(:name => 'Excel 2010 comments') do |sheet|
  sheet.add_row ['Cell with visible comment']
  sheet.add_row
  sheet.add_row
  sheet.add_row ['Cell with hidden comment']

  sheet.add_comment :ref => 'A1', :author => 'XXX', :text => 'Visibile'
  sheet.add_comment :ref => 'A4', :author => 'XXX', :text => 'Hidden', :visible => false
end
p.serialize('excel_2010_comment_test.xlsx')
