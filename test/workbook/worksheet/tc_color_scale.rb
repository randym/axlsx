require 'tc_helper.rb'

class TestColorScale < Test::Unit::TestCase
  def setup
    @color_scale = Axlsx::ColorScale.new
  end

  def test_three_tone
    color_scale = Axlsx::ColorScale.three_tone
    assert_equal 3, color_scale.value_objects.size
    assert_equal 3, color_scale.colors.size
  end

  def test_two_tone
    color_scale = Axlsx::ColorScale.two_tone
    assert_equal 2, color_scale.value_objects.size
    assert_equal 2, color_scale.colors.size
  end
  def test_default_cfvo
    first = Axlsx::ColorScale.default_cfvos.first
    second = Axlsx::ColorScale.default_cfvos.last
    assert_equal 'FFFF7128', first[:color]
    assert_equal :min,first[:type]
    assert_equal 0, first[:val]

    assert_equal 'FFFFEF9C', second[:color]
    assert_equal :max, second[:type]
    assert_equal 0, second[:val]
  end

  def test_partial_default_cfvo_override
    first_def = {:type => :percent, :val => "10.0", :color => 'FF00FF00'}
    color_scale = Axlsx::ColorScale.new(first_def)
    assert_equal color_scale.value_objects.first.val, first_def[:val]
    assert_equal color_scale.value_objects.first.type, first_def[:type]
    assert_equal color_scale.colors.first.rgb, first_def[:color]
  end

  def test_add
    @color_scale.add :type => :max, :val => 5, :color => "FFDEDEDE"
    assert_equal(@color_scale.value_objects.size,3)
    assert_equal(@color_scale.colors.size,3)
  end

  def test_delete_at
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
