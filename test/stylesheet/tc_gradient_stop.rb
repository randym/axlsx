require 'tc_helper.rb'

class TestGradientStop < Test::Unit::TestCase

  def setup
    @item = Axlsx::GradientStop.new(Axlsx::Color.new(:rgb=>"FFFF0000"), 1.0)
  end

  def teardown
  end


  def test_initialiation
    assert_equal(@item.color.rgb, "FFFF0000")
    assert_equal(@item.position, 1.0)
  end

  def test_position
    assert_raise(ArgumentError) { @item.position = -1.1 }
    assert_nothing_raised { @item.position = 0.0 }
    assert_equal(@item.position, 0.0)
  end

  def test_color
    assert_raise(ArgumentError) { @item.color = nil }
    color = Axlsx::Color.new(:rgb=>"FF0000FF")
    @item.color = color
    assert_equal(@item.color.rgb, "FF0000FF")
  end

end
