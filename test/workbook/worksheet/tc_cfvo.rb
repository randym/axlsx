require 'tc_helper.rb'

class TestCfvo < Test::Unit::TestCase
  def setup
    @cfvo = Axlsx::Cfvo.new(:val => "0", :type => :min)
  end

  def test_val
    assert_nothing_raised { @cfvo.val = "abc" }
    assert_equal(@cfvo.val, "abc")
  end

  def test_type
    assert_raise(ArgumentError) { @cfvo.type = :invalid_type }
    assert_nothing_raised { @cfvo.type = :max }
    assert_equal(@cfvo.type, :max)
  end

  def test_gte
    assert_raise(ArgumentError) { @cfvo.gte = :bob }
    assert_equal(@cfvo.gte, true)
    assert_nothing_raised { @cfvo.gte = false }
    assert_equal(@cfvo.gte, false)
  end

  def test_to_xml_string
    doc = Nokogiri::XML.parse(@cfvo.to_xml_string)
    assert doc.xpath(".//cfvo[@type='min'][@val=0][@gte=true]")
  end

end
