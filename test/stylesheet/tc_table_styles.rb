require 'tc_helper.rb'

class TestTableStyles < Test::Unit::TestCase

  def setup
    @item = Axlsx::TableStyles.new
  end

  def teardown
  end

  def test_initialiation
    assert_equal(@item.defaultTableStyle, "TableStyleMedium9")
    assert_equal(@item.defaultPivotStyle, "PivotStyleLight16")
  end

  def test_defaultTableStyle
    assert_raise(ArgumentError) { @item.defaultTableStyle = -1.1 }
    assert_nothing_raised { @item.defaultTableStyle = "anyones guess" }
    assert_equal(@item.defaultTableStyle, "anyones guess")
  end

  def test_defaultPivotStyle
    assert_raise(ArgumentError) { @item.defaultPivotStyle = -1.1 }
    assert_nothing_raised { @item.defaultPivotStyle = "anyones guess" }
    assert_equal(@item.defaultPivotStyle, "anyones guess")
  end

end
