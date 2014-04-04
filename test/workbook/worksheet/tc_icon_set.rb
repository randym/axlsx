require 'tc_helper.rb'

class TestIconSet < Test::Unit::TestCase
  def setup
    @icon_set = Axlsx::IconSet.new
  end

  def test_defaults
    assert_equal @icon_set.iconSet, "3TrafficLights1"
    assert_equal @icon_set.percent, true
    assert_equal @icon_set.reverse, false
    assert_equal @icon_set.showValue, true
  end

  def test_icon_set
    assert_raise(ArgumentError) { @icon_set.iconSet = "invalid_value" }
    assert_nothing_raised { @icon_set.iconSet = "5Rating"}
    assert_equal(@icon_set.iconSet, "5Rating")
  end

  def test_percent
    assert_raise(ArgumentError) { @icon_set.percent = :invalid_type }
    assert_nothing_raised { @icon_set.percent =  false}
    assert_equal(@icon_set.percent, false)
  end

  def test_showValue
    assert_raise(ArgumentError) { @icon_set.showValue = :invalid_type }
    assert_nothing_raised { @icon_set.showValue =  false}
    assert_equal(@icon_set.showValue, false)
  end

  def test_reverse
    assert_raise(ArgumentError) { @icon_set.reverse = :invalid_type }
    assert_nothing_raised { @icon_set.reverse =  false}
    assert_equal(@icon_set.reverse, false)
  end

  def test_to_xml_string
    doc = Nokogiri::XML.parse(@icon_set.to_xml_string)
    assert_equal(doc.xpath(".//iconSet[@iconSet='3TrafficLights1'][@percent=1][@reverse=0][@showValue=1]").size, 1)
    assert_equal(doc.xpath(".//iconSet//cfvo").size, 3)
  end

end
