require 'tc_helper.rb'

class TestGraphicFrame < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
    @chart = @ws.add_chart Axlsx::Chart
    @frame = @chart.graphic_frame
  end

  def teardown
  end

  def test_initialization
    assert(@frame.anchor.is_a?(Axlsx::TwoCellAnchor))
    assert_equal(@frame.chart, @chart)
  end

  def test_rId
    assert_equal(@frame.rId, "rId1")
    chart = @ws.add_chart Axlsx::Chart
    assert_equal(chart.graphic_frame.rId, "rId2")
  end

end
