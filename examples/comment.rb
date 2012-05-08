#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

wb.add_worksheet(:name=>'look mom, comments') do |ws|
  ws.add_row [1,2,3]
  ws.add_comment :text => "now that is just sexy", :author => "randym", :ref => "A1"
end

p.validate.each { |e| puts e.message }
p.serialize 'comments.xlsx'
