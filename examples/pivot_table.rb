#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-

$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'

p = Axlsx::Package.new
wb = p.workbook

# Create some data in a sheet
def month
  %w(Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec).sample
end
def year
  %w(2010 2011 2012).sample
end
def type
  %w(Meat Dairy Beverages Produce).sample
end
def sales
  rand(5000)
end
def region
  %w(East West North South).sample
end

wb.add_worksheet(:name => "Data Sheet") do |sheet|
  sheet.add_row ['Month', 'Year', 'Type', 'Sales', 'Region']
  30.times { sheet.add_row [month, year, type, sales, region] }
  sheet.add_pivot_table 'G4:L17', "A1:E31" do |pivot_table|
    pivot_table.rows = ['Month', 'Year']
    pivot_table.columns = ['Type']
    pivot_table.data = ['Sales']
    pivot_table.pages = ['Region']
  end
end

# Write the excel file
p.serialize("pivot_table.xlsx")
