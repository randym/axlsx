require 'tc_helper.rb'

class TestPictureLocking < Minitest::Unit::TestCase
  def setup
    @item = Axlsx::PictureLocking.new
  end
  def teardown
  end

  def test_initialiation
    assert_equal(@item.instance_values.size, 1)
    assert_equal(@item.noChangeAspect, true)
  end

  def test_noGrp
    assert_raises(ArgumentError) { @item.noGrp = -1 }
    assert_nothing_raised { @item.noGrp = false }
    assert_equal(@item.noGrp, false )
  end

  def test_noRot
    assert_raises(ArgumentError) { @item.noRot = -1 }
    assert_nothing_raised { @item.noRot = false }
    assert_equal(@item.noRot, false )
  end

  def test_noChangeAspect
    assert_raises(ArgumentError) { @item.noChangeAspect = -1 }
    assert_nothing_raised { @item.noChangeAspect = false }
    assert_equal(@item.noChangeAspect, false )
  end

  def test_noMove
    assert_raises(ArgumentError) { @item.noMove = -1 }
    assert_nothing_raised { @item.noMove = false }
    assert_equal(@item.noMove, false )
  end

  def test_noResize
    assert_raises(ArgumentError) { @item.noResize = -1 }
    assert_nothing_raised { @item.noResize = false }
    assert_equal(@item.noResize, false )
  end

  def test_noEditPoints
    assert_raises(ArgumentError) { @item.noEditPoints = -1 }
    assert_nothing_raised { @item.noEditPoints = false }
    assert_equal(@item.noEditPoints, false )
  end

  def test_noAdjustHandles
    assert_raises(ArgumentError) { @item.noAdjustHandles = -1 }
    assert_nothing_raised { @item.noAdjustHandles = false }
    assert_equal(@item.noAdjustHandles, false )
  end

  def test_noChangeArrowheads
    assert_raises(ArgumentError) { @item.noChangeArrowheads = -1 }
    assert_nothing_raised { @item.noChangeArrowheads = false }
    assert_equal(@item.noChangeArrowheads, false )
  end

  def test_noChangeShapeType
    assert_raises(ArgumentError) { @item.noChangeShapeType = -1 }
    assert_nothing_raised { @item.noChangeShapeType = false }
    assert_equal(@item.noChangeShapeType, false )
  end




end
