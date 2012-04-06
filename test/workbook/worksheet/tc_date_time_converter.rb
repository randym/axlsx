# -*- coding: utf-8 -*-
require 'tc_helper.rb'

class TestDateTimeConverter < Test::Unit::TestCase
  def setup
    @margin_of_error = 0.000_001
    @extended_time_range = begin
      Time.parse "1893-08-05"
      Time.parse "9999-12-31T23:59:59Z"
      true
    rescue
      false
    end
  end

  def test_date_to_serial_1900
    Axlsx::Workbook.date1904 = false
    tests = if @extended_time_range
              { # examples taken straight from the spec
                "1893-08-05" => -2338.0,
                "1900-01-01" => 2.0,
                "1910-02-03" => 3687.0,
                "2006-02-01" => 38749.0,
                "9999-12-31" => 2958465.0
              }
            else
              { # examples taken inside the possible values
                "1970-01-01" => 25569.0, # Unix epoch
                "1970-01-02" => 25570.0,
                "2006-02-01" => 38749.0,
                "2038-01-19" => 50424.0, # max date using signed timestamp in 32bit
              }
            end
    tests.each do |date_string, expected|
      serial = Axlsx::DateTimeConverter::date_to_serial Date.parse(date_string)
      assert_equal expected, serial
    end
  end

  def test_date_to_serial_1904
    Axlsx::Workbook.date1904 = true
    tests = if @extended_time_range
              { # examples taken straight from the spec
                "1893-08-05" => -3800.0,
                "1904-01-01" => 0.0,
                "1910-02-03" => 2225.0,
                "2006-02-01" => 37287.0,
                "9999-12-31" => 2957003.0
              }
            else
              { # examples taken inside the possible values
                "1970-01-01" => 24107.0, # Unix epoch
                "1970-01-02" => 24108.0,
                "2006-02-01" => 37287.0,
                "2038-01-19" => 48962.0, # max date using signed timestamp in 32bit
              }
            end
    tests.each do |date_string, expected|
      serial = Axlsx::DateTimeConverter::date_to_serial Date.parse(date_string)
      assert_equal expected, serial
    end
  end

  def test_time_to_serial_1900
    Axlsx::Workbook.date1904 = false
    tests = if @extended_time_range
             { # examples taken straight from the spec
               "1893-08-05T00:00:01Z" => -2337.999989,
               "1899-12-28T18:00:00Z" => -1.25,
               "1910-02-03T10:05:54Z" => 3687.4207639,
               "1900-01-01T12:00:00Z" => 2.5, # wrongly indicated as 1.5 in the spec!
               "9999-12-31T23:59:59Z" => 2958465.9999884
             }
           else
             { # examples taken inside the possible values
               "1970-01-01T00:00:00Z" => 25569.0, # Unix epoch
               "1970-01-01T12:00:00Z" => 25569.5,
               "2000-01-01T00:00:00Z" => 36526.0,
               "2038-01-19T03:14:07Z" => 50424.134803, # max signed timestamp in 32bit
             }
           end
    tests.each do |time_string, expected|
      serial = Axlsx::DateTimeConverter::time_to_serial Time.parse(time_string)
      assert_in_delta expected, serial, @margin_of_error
    end
  end

  def test_time_to_serial_1904
    Axlsx::Workbook.date1904 = true
      # ruby 1.8.7 cannot parse dates prior to epoch. see http://ruby-doc.org/core-1.8.7/Time.html

    tests = if @extended_time_range
             { # examples taken straight from the spec
               "1893-08-05T00:00:01Z" => -3799.999989,
               "1910-02-03T10:05:54Z" => 2225.4207639,
               "1904-01-01T12:00:00Z" => 0.5000000,
               "9999-12-31T23:59:59Z" => 2957003.9999884
             }
           else
             { # examples taken inside the possible values
               "1970-01-01T00:00:00Z" => 24107.0, # Unix epoch
               "1970-01-01T12:00:00Z" => 24107.5,
               "2000-01-01T00:00:00Z" => 35064.0,
               "2038-01-19T03:14:07Z" => 48962.134803, # max signed timestamp in 32bit
             }
           end
    tests.each do |time_string, expected|
      serial = Axlsx::DateTimeConverter::time_to_serial Time.parse(time_string)
      assert_in_delta expected, serial, @margin_of_error
    end
  end

  def test_timezone

    utc = Time.utc 2012 # January 1st, 2012 at 0:00 UTC

    # JRuby makes no assumption on time zone. randym
    #local = begin
    #  Time.new 2012, 1, 1, 1, 0, 0, 3600 # January 1st, 2012 at 1:00 GMT+1
    #rescue ArgumentError
    #  Time.parse "2012-01-01 01:00:00 +0100"
    #end

    local = Time.parse "2012-01-01 01:00:00 +0100"

    assert_equal local, utc
    assert_equal Axlsx::DateTimeConverter::time_to_serial(local), Axlsx::DateTimeConverter::time_to_serial(utc)
    Axlsx::Workbook.date1904 = true
    assert_equal Axlsx::DateTimeConverter::time_to_serial(local), Axlsx::DateTimeConverter::time_to_serial(utc)
  end

end
