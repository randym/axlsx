require 'tc_helper.rb'

class TestScaling < Test::Unit::TestCase
  def setup
    @scaling = Axlsx::Scaling.new
  end

  def teardown
  end

  def test_initialization
    assert(@scaling.orientation == :minMax)
  end

  def test_logBase
    assert_raise(ArgumentError) { @scaling.logBase = 1}
    assert_nothing_raised {@scaling.logBase = 10}
  end

  def test_orientation
    assert_raise(ArgumentError) { @scaling.orientation = "1"}
    assert_nothing_raised {@scaling.orientation = :maxMin}
  end


  def test_max
    assert_raise(ArgumentError) { @scaling.max = 1}
    assert_nothing_raised {@scaling.max = 10.5}
  end

  def test_min
    assert_raise(ArgumentError) { @scaling.min = 1}
    assert_nothing_raised {@scaling.min = 10.5}
  end

end
