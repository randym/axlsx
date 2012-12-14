#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

#```ruby
require 'axlsx'

p = Axlsx::Package.new
p.use_shared_strings = true
s = p.workbook.add_worksheet(:name => "Formula test")
s.add_row [1, 2, 3]
s.add_row %w(a b c)
s.add_row ["=SUM(A1:C1)"], :formula_values => [6]
p.serialize "ios_preview.xlsx"
