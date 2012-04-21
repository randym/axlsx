require 'tc_helper.rb'

class TestColorScale < Test::Unit::TestCase
  def setup
    @color_scale = Axlsx::ColorScale.new
  end

  def test_add
    @color_scale.add :type => :max, :val => 5, :color => "FFDEDEDE"
    assert_equal(@color_scale.value_objects.size,3)
    assert_equal(@color_scale.colors.size,3)
  end

  def test_delete_at
    assert_raise(ArgumentError, "minimum two are protected") { @color_scale.delete_at 0 }
    assert_raise(ArgumentError, "minimum two are protected") { @color_scale.delete_at 1 }
    @color_scale.add :type => :max, :val => 5, :color => "FFDEDEDE"
    assert_nothing_raised {@color_scale.delete_at 2}
    assert_equal(@color_scale.value_objects.size,2)
    assert_equal(@color_scale.colors.size,2)
  end

  def test_to_xml_string
    doc = Nokogiri::XML.parse(@color_scale.to_xml_string)
    assert_equal(doc.xpath(".//colorScale//cfvo").size, 2)
    assert_equal(doc.xpath(".//colorScale//color").size, 2)
  end

end
