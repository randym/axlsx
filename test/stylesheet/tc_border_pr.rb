require 'tc_helper.rb'

class TestBorderPr < Minitest::Unit::TestCase
  def setup
    @bpr = Axlsx::BorderPr.new
  end
  def teardown
  end
  def test_initialiation
    assert_equal(@bpr.color, nil)
    assert_equal(@bpr.style, nil)
    assert_equal(@bpr.name, nil)
  end

  def test_color
    assert_raises(ArgumentError) { @bpr.color = :red }
    assert_nothing_raised { @bpr.color = Axlsx::Color.new :rgb=>"FF000000" }
    assert(@bpr.color.is_a?(Axlsx::Color))
  end

  def test_style
    assert_raises(ArgumentError) { @bpr.style = :red }
    assert_nothing_raised { @bpr.style = :thin }
    assert_equal(@bpr.style, :thin)
  end

  def test_name
    assert_raises(ArgumentError) { @bpr.name = :red }
    assert_nothing_raised { @bpr.name = :top }
    assert_equal(@bpr.name, :top)
  end
end
