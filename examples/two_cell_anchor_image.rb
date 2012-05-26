#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'

p = Axlsx::Package.new
src = "#{File.dirname(__FILE__)}/image1.png"
p.workbook.add_worksheet(:name => 'double_anchor') do |ws|
  ws.add_image(:image_src => src, :start_at => [0,0], :end_at => [2,4])
end
p.serialize('two_cell_anchor_image.xlsx')
