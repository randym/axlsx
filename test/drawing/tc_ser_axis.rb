require 'test/unit'
require 'axlsx.rb'

class TestSerAxis < Test::Unit::TestCase
  def setup    
    @axis = Axlsx::SerAxis.new 12345, 54321
  end
  def teardown
  end

  def test_tickLblSkip
    assert_raise(ArgumentError, "requires valid tickLblSkip") { @axis.tickLblSkip = :my_eyes }
    assert_nothing_raised("accepts valid tickLblSkip") { @axis.tickLblSkip = false }
    assert_equal(@axis.tickLblSkip, false)
  end

  def test_tickMarkSkip
    assert_raise(ArgumentError, "requires valid tickMarkSkip") { @axis.tickMarkSkip = :my_eyes }
    assert_nothing_raised("accepts valid tickMarkSkip") { @axis.tickMarkSkip = false }
    assert_equal(@axis.tickMarkSkip, false)
  end

end
