require 'tc_helper.rb'

class TestPatternFill < Test::Unit::TestCase

  def setup
    @item = Axlsx::PatternFill.new
  end

  def teardown
  end


  def test_initialiation
    assert_equal(@item.patternType, :none)
    assert_equal(@item.bgColor, nil)
    assert_equal(@item.fgColor, nil)
  end

  def test_bgColor
    assert_raise(ArgumentError) { @item.bgColor = -1.1 }
    assert_nothing_raised { @item.bgColor = Axlsx::Color.new }
    assert_equal(@item.bgColor.rgb, "FF000000")
  end

  def test_fgColor
    assert_raise(ArgumentError) { @item.fgColor = -1.1 }
    assert_nothing_raised { @item.fgColor = Axlsx::Color.new }
    assert_equal(@item.fgColor.rgb, "FF000000")
  end

  def test_pattern_type
    assert_raise(ArgumentError) { @item.patternType = -1.1 }
    assert_nothing_raised { @item.patternType = :lightUp }
    assert_equal(@item.patternType, :lightUp)
  end
 
  def test_to_xml_string
    @item = Axlsx::PatternFill.new :bgColor => Axlsx::Color.new(:rgb => "FF0000"), :fgColor => Axlsx::Color.new(:rgb=>"00FF00") 
    doc = Nokogiri::XML(@item.to_xml_string)
    assert(doc.xpath('//color[@rgb="FFFF0000"]'))
    assert(doc.xpath('//color[@rgb="FF00FF00"]'))
  end
end
