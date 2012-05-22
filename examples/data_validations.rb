#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'

p = Axlsx::Package.new
p.workbook.add_worksheet do |ws| 
  data_validation_1 = { :type => :whole, :formula1 => '4', :formula2 => '10'}

  ws.add_data_validation("A1", data_validation_1)
end
p.serialize 'data_validations.xlsx'
