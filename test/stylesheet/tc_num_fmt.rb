require 'tc_helper.rb'

class TestNumFmt < Test::Unit::TestCase

  def setup
    @item = Axlsx::NumFmt.new
  end

  def teardown
  end


  def test_initialiation
    assert_equal(@item.numFmtId, 0)
    assert_equal(@item.formatCode, "")
  end

  def test_numFmtId
    assert_raise(ArgumentError) { @item.numFmtId = -1.1 }
    assert_nothing_raised { @item.numFmtId = 2 }
    assert_equal(@item.numFmtId, 2)
  end

  def test_fomatCode
    assert_raise(ArgumentError) { @item.formatCode = -1.1 }
    assert_nothing_raised { @item.formatCode = "0" }
    assert_equal(@item.formatCode, "0")
  end

end
