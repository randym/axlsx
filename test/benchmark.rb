#!/usr/bin/env ruby -s
$:.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'csv'
require 'benchmark'
Axlsx::trust_input = true
row = []
input = (32..126).to_a.pack('U*').chars.to_a
20.times { row << input.shuffle.join}
times = 3000
Benchmark.bmbm(30) do |x|

  x.report('axlsx_noautowidth') {
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    p.use_autowidth = false
    p.serialize("example_noautowidth.xlsx")
  }

  x.report('axlsx') {
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    p.serialize("example_autowidth.xlsx")
  }

  x.report('axlsx_shared') {
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    p.use_shared_strings = true
    p.serialize("example_shared.xlsx")
  }

  x.report('axlsx_stream') {
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
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
File.delete("example.csv", "example_streamed.xlsx", "example_shared.xlsx", "example_autowidth.xlsx", "example_noautowidth.xlsx")
