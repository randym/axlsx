require 'tc_helper.rb'

class TestMarker < Test::Unit::TestCase
  def setup
    @marker = Axlsx::Marker.new
  end

  def teardown
  end

  def test_initialization
    assert(@marker.col == 0)
    assert(@marker.colOff == 0)
    assert(@marker.row == 0)
    assert(@marker.rowOff == 0)
  end

  def test_col
    assert_raise(ArgumentError) { @marker.col = -1}
    assert_nothing_raised {@marker.col = 10}
  end

  def test_colOff
    assert_raise(ArgumentError) { @marker.colOff = "1"}
    assert_nothing_raised {@marker.colOff = -10}
  end

  def test_row
    assert_raise(ArgumentError) { @marker.row = -1}
    assert_nothing_raised {@marker.row = 10}
  end

  def test_rowOff
    assert_raise(ArgumentError) { @marker.rowOff = "1"}
    assert_nothing_raised {@marker.rowOff = -10}
  end

  def test_coord
    @marker.coord 5, 10
    assert_equal(@marker.col, 5)
    assert_equal(@marker.row, 10)
  end

end
