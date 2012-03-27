require 'tc_helper.rb'

class TestBorder < Test::Unit::TestCase
  def setup
    @b = Axlsx::Border.new
  end
  def teardown
  end
  def test_initialiation
    assert_equal(@b.diagonalUp, nil)
    assert_equal(@b.diagonalDown, nil)
    assert_equal(@b.outline, nil)
    assert(@b.prs.is_a?(Axlsx::SimpleTypedList))
  end

  def test_diagonalUp
    assert_raise(ArgumentError) { @b.diagonalUp = :red }
    assert_nothing_raised { @b.diagonalUp = true }
    assert_equal(@b.diagonalUp, true )
  end

  def test_diagonalDown
    assert_raise(ArgumentError) { @b.diagonalDown = :red }
    assert_nothing_raised { @b.diagonalDown = true }
    assert_equal(@b.diagonalDown, true )
  end

  def test_outline
    assert_raise(ArgumentError) { @b.outline = :red }
    assert_nothing_raised { @b.outline = true }
    assert_equal(@b.outline, true )
  end

  def test_prs
    assert_nothing_raised { @b.prs << Axlsx::BorderPr.new(:name=>:top, :style=>:thin, :color => Axlsx::Color.new(:rgb=>"FF0000FF")) }
  end
end
