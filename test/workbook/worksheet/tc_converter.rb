require 'test/unit'
require 'axlsx.rb'

class TestConverter < Test::Unit::TestCase
  def setup
    @converter = Axlsx::Converter.new
    @margin_of_error = 0.000_001
  end

  def test_date_to_serial_1900
    { # examples taken straight from the spec
      "1893-08-05" => -2338.0,
      "1900-01-01" => 2.0,
      "1910-02-03" => 3687.0,
      "2006-02-01" => 38749.0,
      "9999-12-31" => 2958465.0,
    }.each do |date_string, expected|
      serial = @converter.date_to_serial Date.parse(date_string)
      assert_equal serial, expected
    end
  end

  def test_date_to_serial_1904
    { # examples taken straight from the spec
      "1893-08-05" => -3800.0,
      "1904-01-01" => 0.0,
      "1910-02-03" => 2225.0,
      "2006-02-01" => 37287.0,
      "9999-12-31" => 2957003.0,
    }.each do |date_string, expected|
      serial = @converter.date_to_serial Date.parse(date_string), true
      assert_equal serial, expected
    end
  end

  def test_time_to_serial_1900
    { # examples taken straight from the spec
      "1893-08-05T00:00:01Z" => -2337.999989,
      "1899-12-28T18:00:00Z" => -1.25,
      "1910-02-03T10:05:54Z" => 3687.4207639,
      "1900-01-01T12:00:00Z" => 2.5, # wrongly indicated as 1.5 in the spec!
      "9999-12-31T23:59:59Z" => 2958465.9999884,
    }.each do |time_string, expected|
      serial = @converter.time_to_serial Time.parse(time_string)
      assert_in_delta serial, expected, @margin_of_error
    end
  end

  def test_time_to_serial_1904
    { # examples taken straight from the spec
      "1893-08-05T00:00:01Z" => -3799.999989,
      "1910-02-03T10:05:54Z" => 2225.4207639,
      "1904-01-01T12:00:00Z" => 0.5000000,
      "9999-12-31T23:59:59Z" => 2957003.9999884,
    }.each do |time_string, expected|
      serial = @converter.time_to_serial Time.parse(time_string), true
      assert_in_delta serial, expected, @margin_of_error
    end
  end

  def test_timezone
    utc = Time.utc 2012 # January 1st, 2012 at 0:00 UTC
    local = Time.new 2012, 1, 1, 1, 0, 0, 3600 # January 1st, 2012 at 1:00 GMT+1
    assert_equal local, utc
    assert_equal @converter.time_to_serial(local), @converter.time_to_serial(utc)
    assert_equal @converter.time_to_serial(local, true), @converter.time_to_serial(utc, true)
  end

end
