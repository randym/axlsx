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
    assert_equal @ws.drawing.relationships.for(@chart).Id, @frame.rId
  end

  def test_to_xml_has_correct_rId
    doc = Nokogiri::XML(@frame.to_xml_string)
    assert_equal @frame.rId, doc.xpath("//c:chart", doc.collect_namespaces).first["r:id"]
  end
end
