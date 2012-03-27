require 'tc_helper.rb'

class TestFill < Test::Unit::TestCase

  def setup
    @item = Axlsx::Fill.new Axlsx::PatternFill.new
  end

  def teardown
  end

  def test_initialiation
    assert(@item.fill_type.is_a?(Axlsx::PatternFill))
    assert_raise(ArgumentError) { Axlsx::Fill.new }
    assert_nothing_raised { Axlsx::Fill.new(Axlsx::GradientFill.new) }
  end

end
