require 'tc_helper.rb'

class TestTwoCellAnchor < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
    @ws.add_row ["one", 1, Time.now]
    chart = @ws.add_chart Axlsx::Bar3DChart
    @anchor = chart.graphic_frame.anchor
  end

  def test_initialization
    assert(@anchor.from.col == 0)
    assert(@anchor.from.row == 0)
    assert(@anchor.to.col == 5)
    assert(@anchor.to.row == 10)
  end

  def test_index
    assert_equal(@anchor.index, @anchor.drawing.anchors.index(@anchor))
  end

  def test_options
    assert_raise(ArgumentError, 'invalid start_at') { @ws.add_chart Axlsx::Chart, :start_at=>"1" }
    assert_raise(ArgumentError, 'invalid end_at') { @ws.add_chart Axlsx::Chart, :start_at=>[1,2], :end_at => ["a", 4] }
    # this is actually raised in the graphic frame
    assert_raise(ArgumentError, 'invalid Chart') { @ws.add_chart Axlsx::TwoCellAnchor }
    a = @ws.add_chart Axlsx::Chart, :start_at => [15, 35], :end_at => [90, 45]
    assert_equal(a.graphic_frame.anchor.from.col, 15)
    assert_equal(a.graphic_frame.anchor.from.row, 35)
    assert_equal(a.graphic_frame.anchor.to.col, 90)
    assert_equal(a.graphic_frame.anchor.to.row, 45)
  end

end
