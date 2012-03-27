require 'tc_helper.rb'

class TestValAxis < Test::Unit::TestCase
  def setup
    @axis = Axlsx::ValAxis.new 12345, 54321
  end
  def teardown
  end

  def test_initialization
    assert_equal(@axis.crossBetween, :between, "axis crossBetween default incorrect")
  end

  def test_options
    a = Axlsx::ValAxis.new 2345, 4321, :crossBetween => :midCat
    assert_equal(a.crossBetween, :midCat)
  end

  def test_crossBetween
    assert_raise(ArgumentError, "requires valid crossBetween") { @axis.crossBetween = :my_eyes }
    assert_nothing_raised("accepts valid crossBetween") { @axis.crossBetween = :midCat }
  end

end
