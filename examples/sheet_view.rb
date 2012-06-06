#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'

p = Axlsx::Package.new
ws = p.workbook.add_worksheet :name => "Sheetview - Split"
ws.sheet_view do |vs|
  vs.pane do |p|
    p.active_pane = :top_right
    p.state = :split
    p.x_split = 11080
    p.y_split = 5000
    p.top_left_cell = 'C44'
  end
  
  vs.add_selection(:top_left, { :active_cell => 'A2', :sqref => 'A2' })
  vs.add_selection(:top_right, { :active_cell => 'I10', :sqref => 'I10' })
  vs.add_selection(:bottom_left, { :active_cell => 'E55', :sqref => 'E55' })
  vs.add_selection(:bottom_right, { :active_cell => 'I57', :sqref => 'I57' })
end


ws = p.workbook.add_worksheet :name => "Sheetview - Frozen"
ws.sheet_view do |vs|
  vs.pane do |p|
    p.state = :frozen
    p.x_split = 3
    p.y_split = 4
  end
end


p.serialize 'sheet_view.xlsx'