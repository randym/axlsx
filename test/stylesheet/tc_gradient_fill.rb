require 'test/unit'
require 'axlsx.rb'

class TestGradientFill < Test::Unit::TestCase

  def setup
    @item = Axlsx::GradientFill.new
  end

  def teardown
  end


  def test_initialiation
    assert_equal(@item.type, :linear)
    assert_equal(@item.degree, nil)
    assert_equal(@item.left, nil)
    assert_equal(@item.right, nil)
    assert_equal(@item.top, nil)
    assert_equal(@item.bottom, nil)
    assert(@item.stop.is_a?(Axlsx::SimpleTypedList))
  end

  def test_type
    assert_raise(ArgumentError) { @item.type = 7 }
    assert_nothing_raised { @item.type = :path }
    assert_equal(@item.type, :path)
  end

  def test_degree
    assert_raise(ArgumentError) { @item.degree = -7 }
    assert_nothing_raised { @item.degree = 5.0 }
    assert_equal(@item.degree, 5.0)
  end

  def test_left
    assert_raise(ArgumentError) { @item.left = -1.1 }
    assert_nothing_raised { @item.left = 1.0 }
    assert_equal(@item.left, 1.0)
  end

  def test_right
    assert_raise(ArgumentError) { @item.right = -1.1 }
    assert_nothing_raised { @item.right = 0.5 }
    assert_equal(@item.right, 0.5)
  end

  def test_top
    assert_raise(ArgumentError) { @item.top = -1.1 }
    assert_nothing_raised { @item.top = 1.0 }
    assert_equal(@item.top, 1.0)
  end

  def test_bottom
    assert_raise(ArgumentError) { @item.bottom = -1.1 }
    assert_nothing_raised { @item.bottom = 0.0 }
    assert_equal(@item.bottom, 0.0)
  end

  def test_stop
    @item.stop << Axlsx::GradientStop.new(Axlsx::Color.new(:rgb=>"00000000"), 0.5)
    assert(@item.stop.size == 1)
    assert(@item.stop.last.is_a?(Axlsx::GradientStop))
  end
end
