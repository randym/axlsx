#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'

p = Axlsx::Package.new
p.workbook.add_worksheet do |ws| 
  ws.add_data_validation("A10", { 
    :type => :whole, 
    :operator => :between, 
    :formula1 => '5', 
    :formula2 => '10', 
    :showErrorMessage => true, 
    :errorTitle => 'Wrong input', 
    :error => 'Only values between 5 and 10', 
    :errorStyle => :information, 
    :showInputMessage => true, 
    :promptTitle => 'Be carful!', 
    :prompt => 'Only values between 5 and 10'})
  
  ws.add_data_validation("B10", { 
    :type => :textLength, 
    :operator => :greaterThan, 
    :formula1 => '10', 
    :showErrorMessage => true, 
    :errorTitle => 'Text is too long', 
    :error => 'Max text length is 10 characters', 
    :errorStyle => :stop, 
    :showInputMessage => true, 
    :promptTitle => 'Text length', 
    :prompt => 'Max text length is 10 characters'})
  
  8.times do |i|
    ws.add_row [nil, nil, i*2]
  end
  
  ws.add_data_validation("C10", { 
    :type => :list, 
    :formula1 => 'C1:C8', 
    :showDropDown => false,
    :showErrorMessage => true, 
    :errorTitle => '', 
    :error => 'Only values from C1:C8', 
    :errorStyle => :stop, 
    :showInputMessage => true, 
    :promptTitle => '', 
    :prompt => 'Only values from C1:C8'})
end

p.serialize 'data_validation.xlsx'