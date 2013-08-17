require 'tc_helper.rb'

class TestColor < Test::Unit::TestCase

  def setup
    @item = Axlsx::Color.new
  end

  def teardown
  end

  def test_initialiation
    assert_equal(@item.auto, nil)
    assert_equal(@item.rgb, "FF000000")
    assert_equal(@item.tint, nil)
  end

  def test_auto
    assert_raise(ArgumentError) { @item.auto = -1 }
    assert_nothing_raised { @item.auto = true }
    assert_equal(@item.auto, true )
  end

  def test_rgb
    assert_raise(ArgumentError) { @item.rgb = -1 }
    assert_nothing_raised { @item.rgb = "FF00FF00" }
    assert_equal(@item.rgb, "FF00FF00" )
  end

  def test_rgb_writer_doesnt_mutate_its_argument
    my_rgb = 'ff00ff00'
    @item.rgb = my_rgb
    assert_equal 'ff00ff00', my_rgb
  end

  def test_tint
    assert_raise(ArgumentError) { @item.tint = -1 }
    assert_nothing_raised { @item.tint = -1.0 }
    assert_equal(@item.tint, -1.0 )
  end


end
