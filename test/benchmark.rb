#!/usr/bin/env ruby -s
# -*- coding: utf-8 -*-
$:.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'csv'

require 'benchmark'
row = []
input = (32..126).to_a.pack('U*').chars.to_a
20.times { row << input.shuffle.join}
times = 1000
Benchmark.bm(100) do |x|
  # No Autowidth
  x.report('axlsx_noautowidth') {

    p = Axlsx::Package.new
    p.use_autowidth = false
    wb = p.workbook

    #A Simple Workbook

    wb.add_worksheet do |sheet|
      times.times do
        sheet << row
      end
    end
    p.serialize("example.xlsx")
  }

  x.report('axlsx') {
    p = Axlsx::Package.new
    wb = p.workbook

    #A Simple Workbook

    wb.add_worksheet do |sheet|
      times.times do
        sheet << row
      end
    end
    p.serialize("example.xlsx")
  }

  x.report('axlsx_shared') {
    p = Axlsx::Package.new
    wb = p.workbook

    #A Simple Workbook

    wb.add_worksheet do |sheet|
      times.times do
        sheet << row
      end
    end
    p.use_shared_strings = true
    p.serialize("example.xlsx")
  }

  x.report('axlsx_stream') {
    p = Axlsx::Package.new
    wb = p.workbook

    #A Simple Workbook

    wb.add_worksheet do |sheet|
      times.times do
        sheet << row
      end
    end

    s = p.to_stream()
    File.open('example_streamed.xlsx', 'w') { |f| f.write(s.read) }
  }
  x.report('csv') {
    CSV.open("example.csv", "wb") do |csv|
      times.times do
        csv << row
      end
    end
  }
end
