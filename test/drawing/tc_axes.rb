require 'tc_helper.rb'

class TestAxes < Test::Unit::TestCase
  def test_constructor_requires_cat_axis_first
    assert_raise(ArgumentError) { Axlsx::Axes.new(:val_axis => Axlsx::ValAxis, :cat_axis => Axlsx::CatAxis) }
    assert_nothing_raised { Axlsx::Axes.new(:cat_axis => Axlsx::CatAxis, :val_axis => Axlsx::ValAxis) }
  end
end