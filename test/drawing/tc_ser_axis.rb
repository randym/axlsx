require 'tc_helper.rb'

class TestSerAxis < Test::Unit::TestCase
  def setup
    @axis = Axlsx::SerAxis.new
  end

  def teardown
  end

  def test_options
    a = Axlsx::SerAxis.new(:tick_lbl_skip => 9, :tick_mark_skip => 7)
    assert_equal(a.tick_lbl_skip, 9)
    assert_equal(a.tick_mark_skip, 7)
  end


  def test_tick_lbl_skip
    assert_raise(ArgumentError, "requires valid tick_lbl_skip") { @axis.tick_lbl_skip = -1 }
    assert_nothing_raised("accepts valid tick_lbl_skip") { @axis.tick_lbl_skip = 1 }
    assert_equal(@axis.tick_lbl_skip, 1)
  end


  def test_tick_mark_skip
    assert_raise(ArgumentError, "requires valid tick_mark_skip") { @axis.tick_mark_skip = :my_eyes }
    assert_nothing_raised("accepts valid tick_mark_skip") { @axis.tick_mark_skip = 2 }
    assert_equal(@axis.tick_mark_skip, 2)
  end

end
