require 'tc_helper.rb'

class TestAxis < Test::Unit::TestCase
  def setup
    @axis = Axlsx::Axis.new 12345, 54321, :gridlines => false
  end
  def teardown
  end

  def test_initialization
    assert_equal(@axis.axPos, :b, "axis position default incorrect")
    assert_equal(@axis.tickLblPos, :nextTo, "tick label position default incorrect")
    assert_equal(@axis.tickLblPos, :nextTo, "tick label position default incorrect")
    assert_equal(@axis.crosses, :autoZero, "tick label position default incorrect")
    assert(@axis.scaling.is_a?(Axlsx::Scaling) && @axis.scaling.orientation == :minMax, "scaling default incorrect")
    assert_raise(ArgumentError) { Axlsx::Axis.new -1234, 'abcd' }
  end

  def test_axis_position
    assert_raise(ArgumentError, "requires valid axis position") { @axis.axPos = :nowhere }
    assert_nothing_raised("accepts valid axis position") { @axis.axPos = :r }
  end

  def test_label_rotation
    assert_raise(ArgumentError, "requires valid angle") { @axis.label_rotation = :nowhere }
    assert_raise(ArgumentError, "requires valid angle") { @axis.label_rotation = 91 }
    assert_raise(ArgumentError, "requires valid angle") { @axis.label_rotation = -91 }
    assert_nothing_raised("accepts valid angle") { @axis.label_rotation = 45 }
    assert_equal(@axis.label_rotation, 45 * 60000)
  end

  def test_tick_label_position
    assert_raise(ArgumentError, "requires valid tick label position") { @axis.tickLblPos = :nowhere }
    assert_nothing_raised("accepts valid tick label position") { @axis.tickLblPos = :high }
  end

  def test_format_code
    assert_raise(ArgumentError, "requires valid format code") { @axis.format_code = 1 }
    assert_nothing_raised("accepts valid format code") { @axis.tickLblPos = :high }
  end

  def test_crosses
    assert_raise(ArgumentError, "requires valid crosses") { @axis.crosses = 1 }
    assert_nothing_raised("accepts valid crosses") { @axis.crosses = :min }
  end

  def test_gridlines
    assert_raise(ArgumentError, "requires valid gridlines") { @axis.gridlines = 'alice' }
    assert_nothing_raised("accepts valid crosses") { @axis.gridlines = false }
  end
  
  def test_to_xml_string
    str  = '<?xml version="1.0" encoding="UTF-8"?>'
    str << '<c:chartSpace xmlns:c="' << Axlsx::XML_NS_C << '" xmlns:a="' << Axlsx::XML_NS_A << '">'
    doc = Nokogiri::XML(@axis.to_xml_string(str)) 
    assert(doc.xpath('//a:noFill'))
    assert(doc.xpath("//c:crosses[@val='#{@crosses.to_s}']"))
    assert(doc.xpath("//c:crossAx[@val='#{@crossAx.to_s}']"))
    assert(doc.xpath("//a:bodyPr[@rot='#{@label_rotation.to_s}']"))

  end
end
