require 'tc_helper.rb'

class TestTableStyle < Test::Unit::TestCase

  def setup
    @item = Axlsx::TableStyle.new "fisher"
  end

  def teardown
  end

  def test_initialiation
    assert_equal(@item.name, "fisher")
    assert_equal(@item.pivot, nil)
    assert_equal(@item.table, nil)
  end

  def test_name
    assert_raise(ArgumentError) { @item.name = -1.1 }
    assert_nothing_raised { @item.name = "lovely table style" }
    assert_equal(@item.name, "lovely table style")
  end

  def test_pivot
    assert_raise(ArgumentError) { @item.pivot = -1.1 }
    assert_nothing_raised { @item.pivot = true }
    assert_equal(@item.pivot, true)
  end

  def test_table
    assert_raise(ArgumentError) { @item.table = -1.1 }
    assert_nothing_raised { @item.table = true }
    assert_equal(@item.table, true)
  end

end
