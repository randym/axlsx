require 'tc_helper.rb'

class TestDrawing < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet

  end

  def test_initialization
    assert(@ws.workbook.drawings.empty?)
  end

  def test_add_chart
    chart = @ws.add_chart(Axlsx::Pie3DChart, :title=>"bob", :start_at=>[0,0], :end_at=>[1,1])
    assert(chart.is_a?(Axlsx::Pie3DChart), "must create a chart")
    assert_equal(@ws.workbook.charts.last, chart, "must be added to workbook charts collection")
    assert_equal(@ws.drawing.anchors.last.object.chart, chart, "an anchor has been created and holds a reference to this chart")
    anchor = @ws.drawing.anchors.last
    assert_equal([anchor.from.row, anchor.from.col], [0,0], "options for start at are applied")
    assert_equal([anchor.to.row, anchor.to.col], [1,1], "options for start at are applied")
    assert_equal(chart.title.text, "bob", "option for title is applied")
  end

  def test_add_image
    src = File.dirname(__FILE__) + "/../../examples/image1.jpeg"
    image = @ws.add_image(:image_src => src, :start_at=>[0,0], :width=>600, :height=>400)
    assert(@ws.drawing.anchors.last.is_a?(Axlsx::OneCellAnchor))
    assert(image.is_a?(Axlsx::Pic))
    assert_equal(600, image.width)
    assert_equal(400, image.height)
  end
  def test_add_two_cell_anchor_image
     src = File.dirname(__FILE__) + "/../../examples/image1.jpeg"
     image = @ws.add_image(:image_src => src, :start_at=>[0,0], :end_at => [15,0])
    assert(@ws.drawing.anchors.last.is_a?(Axlsx::TwoCellAnchor))
    assert(image.is_a?(Axlsx::Pic))
  end
  def test_charts
    chart = @ws.add_chart(Axlsx::Pie3DChart, :title=>"bob", :start_at=>[0,0], :end_at=>[1,1])
    assert_equal(@ws.drawing.charts.last, chart, "add chart is returned")
    chart = @ws.add_chart(Axlsx::Pie3DChart, :title=>"nancy", :start_at=>[1,5], :end_at=>[5,10])
    assert_equal(@ws.drawing.charts.last, chart, "add chart is returned")
  end

  def test_pn
    @ws.add_chart(Axlsx::Pie3DChart)
    assert_equal(@ws.drawing.pn, "drawings/drawing1.xml")
  end

  def test_rels_pn
    @ws.add_chart(Axlsx::Pie3DChart)
    assert_equal(@ws.drawing.rels_pn, "drawings/_rels/drawing1.xml.rels")
  end

  def test_rId
    @ws.add_chart(Axlsx::Pie3DChart)
    assert_equal(@ws.drawing.rId, "rId1")
  end

  def test_index
    @ws.add_chart(Axlsx::Pie3DChart)
    assert_equal(@ws.drawing.index, @ws.workbook.drawings.index(@ws.drawing))
  end

  def test_relationships
    chart = @ws.add_chart(Axlsx::Pie3DChart, :title=>"bob", :start_at=>[0,0], :end_at=>[1,1])
    assert_equal(@ws.drawing.relationships.size, 1, "adding a chart adds a relationship")
    chart = @ws.add_chart(Axlsx::Pie3DChart, :title=>"nancy", :start_at=>[1,5], :end_at=>[5,10])
    assert_equal(@ws.drawing.relationships.size, 2, "adding a chart adds a relationship")
  end

  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    @ws.add_chart(Axlsx::Pie3DChart)
    doc = Nokogiri::XML(@ws.drawing.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.empty?, "error free validation")
  end

end
