# encoding: UTF-8
require "date"

module Axlsx
  # The DateTimeConverter class converts both data and time types to their apprpriate excel serializations
  class DateTimeConverter

    # The date_to_serial method converts Date objects to the equivelant excel serialized forms
    # @param [Date] date the date to be serialized
    # @return [Numeric]
    def self.date_to_serial(date)
      epoch = Axlsx::Workbook::date1904 ? Date.new(1904) : Date.new(1899, 12, 30)
      offset_date = date.respond_to?(:utc_offset) ? date + date.utc_offset.seconds : date
      (offset_date - epoch).to_f
    end

    # The time_to_serial methond converts a Time object its excel serialized form.
    # @param [Time] time the time to be serialized
    # @return [Numeric]
    def self.time_to_serial(time)
      # Using hardcoded offsets here as some operating systems will not except
      # a 'negative' offset from the ruby epoch.
      epoch1900 = -2209161600.0 # Time.utc(1899, 12, 30).to_i
      epoch1904 = -2082844800.0 # Time.utc(1904, 1, 1).to_i
      seconds_per_day = 86400.0 # 60*60*24
      epoch = Axlsx::Workbook::date1904 ? epoch1904 : epoch1900
      (time.utc_offset + time.to_f - epoch)/seconds_per_day
    end
  end
end
