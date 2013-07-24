require 'tc_helper.rb'

class TestValAxis < Test::Unit::TestCase
  def setup
    @axis = Axlsx::ValAxis.new
  end
  def teardown
  end

  def test_initialization
    assert_equal(@axis.cross_between, :between, "axis crossBetween default incorrect")
  end

  def test_options
    a = Axlsx::ValAxis.new(:cross_between => :midCat)
    assert_equal(:midCat, a.cross_between)
  end

  def test_crossBetween
    assert_raise(ArgumentError, "requires valid crossBetween") { @axis.cross_between = :my_eyes }
    assert_nothing_raised("accepts valid crossBetween") { @axis.cross_between = :midCat }
  end

end
