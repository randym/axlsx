require 'tc_helper.rb'

class TestSerAxis < Test::Unit::TestCase
  def setup
    @axis = Axlsx::SerAxis.new 12345, 54321
  end
  def teardown
  end

  def test_options
    a = Axlsx::SerAxis.new 12345, 54321, :tickLblSkip => 9, :tickMarkSkip => 7
    assert_equal(a.tickLblSkip, 9)
    assert_equal(a.tickMarkSkip, 7)
  end


  def test_tickLblSkip
    assert_raise(ArgumentError, "requires valid tickLblSkip") { @axis.tickLblSkip = -1 }
    assert_nothing_raised("accepts valid tickLblSkip") { @axis.tickLblSkip = 1 }
    assert_equal(@axis.tickLblSkip, 1)
  end


  def test_tickMarkSkip
    assert_raise(ArgumentError, "requires valid tickMarkSkip") { @axis.tickMarkSkip = :my_eyes }
    assert_nothing_raised("accepts valid tickMarkSkip") { @axis.tickMarkSkip = 2 }
    assert_equal(@axis.tickMarkSkip, 2)
  end

end
