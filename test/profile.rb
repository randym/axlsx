#!/usr/bin/env ruby -s

$:.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'ruby-prof'
row = []
# Taking worst case scenario of all string data
input = (32..126).to_a.pack('U*').chars.to_a
20.times { row << input.shuffle.join}

profile = RubyProf.profile do
  p = Axlsx::Package.new
  p.workbook.add_worksheet do |sheet|
    30.times do
      sheet << row
    end
  end
end
printer = RubyProf::CallTreePrinter.new(profile)
printer.print(File.new('axlsx.qcachegrind', 'w'))
