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

  def test_rId_with_image_and_chart
    image = @ws.add_image :image_src => (File.dirname(__FILE__) + "/../../examples/image1.jpeg"), :start_at => [0,25], :width => 200, :height => 200
    assert_equal(2, image.id)
    assert_equal(1, @chart.index+1)
  end
end
