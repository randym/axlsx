#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

#```ruby
require 'axlsx'
package = Axlsx::Package.new
package.workbook do |workbook|
  workbook.styles do |s|
    gridstyle_border =  s.add_style :border => { :style => :thin, :color =>"FFCDCDCD" }
    workbook.add_worksheet :name => "Custom Borders"  do |sheet|
      sheet.sheet_view.show_grid_lines = false
      sheet.add_row ["with", "grid", "style"], :style => gridstyle_border
      sheet.add_row ["no", "border"]
    end
  end
end
package.serialize 'no_grid_with_borders.xlsx'
