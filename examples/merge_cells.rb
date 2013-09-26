#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

#```ruby
require 'axlsx'
package = Axlsx::Package.new
package.workbook do |workbook|
  workbook.add_worksheet name: 'merged_cells' do |sheet|
    4.times do
      sheet.add_row %w(a b c d e f g)
    end
    sheet.merge_cells "A1:A2"
    sheet.merge_cells "B1:B2"
  end
end
package.serialize 'merged_cells.xlsx'
