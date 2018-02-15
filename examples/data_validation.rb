#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'

p = Axlsx::Package.new
p.workbook.add_worksheet do |ws|

  ws.add_row ["between", "lessThan", "bound list", "raw list"]

  4.times do |i|
    ws.add_row [nil, nil, nil, nil, (i+1) * 2]
  end

  ws.add_data_validation("A2:A5", {
    :type => :whole,
    :operator => :between,
   :formula1 => '5',
    :formula2 => '10',
    :showErrorMessage => true,
    :errorTitle => 'Wrong input',
    :error => 'Only values between 5 and 10',
    :errorStyle => :information,
    :showInputMessage => true,
    :promptTitle => 'Be careful!',
    :prompt => %{We really want a value between 5 and 10,
but it is OK if you want to break the rules.
}})

 ws.add_data_validation("B1:B5", {
   :type => :textLength,
   :operator => :lessThan,
   :formula1 => '10',
   :showErrorMessage => true,
   :errorTitle => 'Text is too long',
   :error => 'Max text length is 10 characters',
   :errorStyle => :stop,
   :showInputMessage => true,
   :promptTitle => 'Text length',
   :prompt => 'Max text length is 10 characters'})

  ws.add_data_validation("C2:C5", {
   :type => :list,
   :formula1 => 'E2:E5',
   :showDropDown => false,
   :showErrorMessage => true,
   :errorTitle => '',
   :error => 'Only values from E2:E5',
   :errorStyle => :stop,
   :showInputMessage => true,
   :promptTitle => '',
   :prompt => 'Only values from E2:E5'})

  ws.add_data_validation("D2:D5", {
    :type => :list,
    :formula1 => '"Red, Orange, NavyBlue"',
    :showDropDown => false,
    :showErrorMessage => true,
    :errorTitle => '',
    :error => 'Please use the dropdown selector to choose the value',
    :errorStyle => :stop,
    :showInputMessage => true,
    :prompt => '&amp; Choose the value from the dropdown'})

end

p.serialize 'data_validation.xlsx'
