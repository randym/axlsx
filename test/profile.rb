#!/usr/bin/env ruby -s
# -*- coding: utf-8 -*-

# Usage:
# > ruby test/profile.rb
# > pprof.rb --gif /tmp/axlsx_noautowidth > /tmp/axlsx_noautowidth.gif
# > open /tmp/axlsx_noautowidth.gif

$:.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'csv'

# require 'benchmark'
require 'perftools'
row = []
input = (32..126).to_a.pack('U*').chars.to_a
20.times { row << input.shuffle.join}
times = 3000

PerfTools::CpuProfiler.start("/tmp/axlsx_noautowidth") do
  p = Axlsx::Package.new
  p.use_autowidth = false
  p.use_shared_strings = true
  wb = p.workbook
  
  #A Simple Workbook
  
  wb.add_worksheet do |sheet|
    times.times do
      sheet << row
    end
  end
  p.serialize("example.xlsx")
end
