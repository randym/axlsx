#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
p = Axlsx::Package.new
wb = p.workbook
wb.add_worksheet name: 'pane' do |sheet|
  sheet.sheet_view.pane do |pane|
    pane.top_left_cell = "B2"
    pane.state = :frozen_split
    pane.y_split = 2
    pane.x_split = 1
    pane.active_pane = :bottom_right
  end
end
p.serialize 'pane.xlsx'
