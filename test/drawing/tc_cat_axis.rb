require 'tc_helper.rb'

class TestCatAxis < Test::Unit::TestCase
  def setup
    @axis = Axlsx::CatAxis.new 12345, 54321
  end
  def teardown
  end

  def test_initialization
    assert_equal(@axis.auto, 1, "axis auto default incorrect")
    assert_equal(@axis.lblAlgn, :ctr, "label align default incorrect")
    assert_equal(@axis.lblOffset, "100", "label offset default incorrect")
  end

  def test_auto
    assert_raise(ArgumentError, "requires valid auto") { @axis.auto = :nowhere }
    assert_nothing_raised("accepts valid auto") { @axis.auto = false }
  end

  def test_lblAlgn
    assert_raise(ArgumentError, "requires valid label alignment") { @axis.lblAlgn = :nowhere }
    assert_nothing_raised("accepts valid label alignment") { @axis.lblAlgn = :r }
  end

  def test_lblOffset
    assert_raise(ArgumentError, "requires valid label offset") { @axis.lblOffset = 'foo' }
    assert_nothing_raised("accepts valid label offset") { @axis.lblOffset = "20" }
  end

end
