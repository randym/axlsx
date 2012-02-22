# encoding: UTF-8
require "date"

module Axlsx
  class Converter
    def date_to_serial(date, date1904=false)
      epoc = date1904 ? Date.new(1904) : Date.new(1899, 12, 30)
      (date-epoc).to_f
    end
      
    def time_to_serial(time, date1904=false)
      # Using hardcoded offsets here as some operating systems will not except
      # a 'negative' offset from the ruby epoc.
      epoc1900 = -2209161600 # Time.utc(1899, 12, 30).to_i
      epoc1904 = -2082844800 # Time.utc(1904, 1, 1).to_i
      seconds_per_day = 86400 # 60*60*24
      epoc = date1904 ? epoc1904 : epoc1900
      (time.to_f - epoc)/seconds_per_day
    end
  end
end
