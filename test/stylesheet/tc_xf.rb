require 'tc_helper.rb'

class TestXf < Test::Unit::TestCase

  def setup
    @item = Axlsx::Xf.new
  end

  def teardown
  end

  def test_initialiation
    assert_equal(@item.alignment, nil)
    assert_equal(@item.protection, nil)
    assert_equal(@item.numFmtId, nil)
    assert_equal(@item.fontId, nil)
    assert_equal(@item.fillId, nil)
    assert_equal(@item.borderId, nil)
    assert_equal(@item.xfId, nil)
    assert_equal(@item.quotePrefix, nil)
    assert_equal(@item.pivotButton, nil)
    assert_equal(@item.applyNumberFormat, nil)
    assert_equal(@item.applyFont, nil)
    assert_equal(@item.applyFill, nil)
    assert_equal(@item.applyBorder, nil)
    assert_equal(@item.applyAlignment, nil)
    assert_equal(@item.applyProtection, nil)
  end

  def test_alignment
    assert_raise(ArgumentError) { @item.alignment = -1.1 }
    assert_nothing_raised { @item.alignment = Axlsx::CellAlignment.new }
    assert(@item.alignment.is_a?(Axlsx::CellAlignment))
  end

  def test_protection
    assert_raise(ArgumentError) { @item.protection = -1.1 }
    assert_nothing_raised { @item.protection = Axlsx::CellProtection.new }
    assert(@item.protection.is_a?(Axlsx::CellProtection))
  end

  def test_numFmtId
    assert_raise(ArgumentError) { @item.numFmtId = -1.1 }
    assert_nothing_raised { @item.numFmtId = 0 }
    assert_equal(@item.numFmtId, 0)
  end

  def test_fillId
    assert_raise(ArgumentError) { @item.fillId = -1.1 }
    assert_nothing_raised { @item.fillId = 0 }
    assert_equal(@item.fillId, 0)
  end

  def test_fontId
    assert_raise(ArgumentError) { @item.fontId = -1.1 }
    assert_nothing_raised { @item.fontId = 0 }
    assert_equal(@item.fontId, 0)
  end

  def test_borderId
    assert_raise(ArgumentError) { @item.borderId = -1.1 }
    assert_nothing_raised { @item.borderId = 0 }
    assert_equal(@item.borderId, 0)
  end

  def test_xfId
    assert_raise(ArgumentError) { @item.xfId = -1.1 }
    assert_nothing_raised { @item.xfId = 0 }
    assert_equal(@item.xfId, 0)
  end

  def test_quotePrefix
    assert_raise(ArgumentError) { @item.quotePrefix = -1.1 }
    assert_nothing_raised { @item.quotePrefix = false }
    assert_equal(@item.quotePrefix, false)
  end

  def test_pivotButton
    assert_raise(ArgumentError) { @item.pivotButton = -1.1 }
    assert_nothing_raised { @item.pivotButton = false }
    assert_equal(@item.pivotButton, false)
  end

  def test_applyNumberFormat
    assert_raise(ArgumentError) { @item.applyNumberFormat = -1.1 }
    assert_nothing_raised { @item.applyNumberFormat = false }
    assert_equal(@item.applyNumberFormat, false)
  end

  def test_applyFont
    assert_raise(ArgumentError) { @item.applyFont = -1.1 }
    assert_nothing_raised { @item.applyFont = false }
    assert_equal(@item.applyFont, false)
  end

  def test_applyFill
    assert_raise(ArgumentError) { @item.applyFill = -1.1 }
    assert_nothing_raised { @item.applyFill = false }
    assert_equal(@item.applyFill, false)
  end

  def test_applyBorder
    assert_raise(ArgumentError) { @item.applyBorder = -1.1 }
    assert_nothing_raised { @item.applyBorder = false }
    assert_equal(@item.applyBorder, false)
  end

  def test_applyAlignment
    assert_raise(ArgumentError) { @item.applyAlignment = -1.1 }
    assert_nothing_raised { @item.applyAlignment = false }
    assert_equal(@item.applyAlignment, false)
  end

  def test_applyProtection
    assert_raise(ArgumentError) { @item.applyProtection = -1.1 }
    assert_nothing_raised { @item.applyProtection = false }
    assert_equal(@item.applyProtection, false)
  end

end
