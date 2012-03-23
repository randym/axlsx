require 'tc_helper.rb'

class TestCellAlignment < Test::Unit::TestCase
  def setup
    @item = Axlsx::CellAlignment.new
  end

  def test_initialiation
    assert_equal(@item.horizontal, nil)
    assert_equal(@item.vertical, nil)
    assert_equal(@item.textRotation, nil)
    assert_equal(@item.wrapText, nil)
    assert_equal(@item.indent, nil)
    assert_equal(@item.relativeIndent, nil)
    assert_equal(@item.justifyLastLine, nil)
    assert_equal(@item.shrinkToFit, nil)
    assert_equal(@item.readingOrder, nil)
    options =  { :horizontal => :left, :vertical => :top, :textRotation => 3,
                 :wrapText => true, :indent => 2, :relativeIndent => 5,
      :justifyLastLine => true, :shrinkToFit => true, :readingOrder => 2 }
    ca = Axlsx::CellAlignment.new options
    options.each do |key, value|
      assert_equal(ca.send(key.to_sym),value)
    end
  end

  def test_horizontal
    assert_raise(ArgumentError) { @item.horizontal = :red }
    assert_nothing_raised { @item.horizontal = :left }
    assert_equal(@item.horizontal, :left )
  end

  def test_vertical
    assert_raise(ArgumentError) { @item.vertical = :red }
    assert_nothing_raised { @item.vertical = :top }
    assert_equal(@item.vertical, :top )
  end

  def test_textRotation
    assert_raise(ArgumentError) { @item.textRotation = -1 }
    assert_nothing_raised { @item.textRotation = 5 }
    assert_equal(@item.textRotation, 5 )
  end

  def test_wrapText
    assert_raise(ArgumentError) { @item.wrapText = -1 }
    assert_nothing_raised { @item.wrapText = false }
    assert_equal(@item.wrapText, false )
  end

  def test_indent
    assert_raise(ArgumentError) { @item.indent = -1 }
    assert_nothing_raised { @item.indent = 5 }
    assert_equal(@item.indent, 5 )
  end

  def test_relativeIndent
    assert_raise(ArgumentError) { @item.relativeIndent = :symbol }
    assert_nothing_raised { @item.relativeIndent = 5 }
    assert_equal(@item.relativeIndent, 5 )
  end

  def test_justifyLastLine
    assert_raise(ArgumentError) { @item.justifyLastLine = -1 }
    assert_nothing_raised { @item.justifyLastLine = true }
    assert_equal(@item.justifyLastLine, true )
  end

  def test_shrinkToFit
    assert_raise(ArgumentError) { @item.shrinkToFit = -1 }
    assert_nothing_raised { @item.shrinkToFit = true }
    assert_equal(@item.shrinkToFit, true )
  end

  def test_readingOrder
    assert_raise(ArgumentError) { @item.readingOrder = -1 }
    assert_nothing_raised { @item.readingOrder = 2 }
    assert_equal(@item.readingOrder, 2 )
  end

end
