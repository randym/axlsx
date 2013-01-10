#!/usr/bin/env ruby -s

# Usage:
# > ruby test/profile.rb
# > pprof.rb --gif /tmp/axlsx > /tmp/axlsx.gif
# > open /tmp/axlsx_noautowidth.gif

$:.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'perftools'
Axlsx.trust_input = true
row = []
# Taking worst case scenario of all string data
input = (32..126).to_a.pack('U*').chars.to_a
20.times { row << input.shuffle.join}
times = 3000

PerfTools::CpuProfiler.start("/tmp/axlsx") do
  p = Axlsx::Package.new
  p.workbook.add_worksheet do |sheet|
    times.times do
      sheet << row
    end
  end
  p.serialize("example.xlsx")
end
