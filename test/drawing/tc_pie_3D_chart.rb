require 'tc_helper.rb'

class TestPie3DChart < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @chart = ws.drawing.add_chart Axlsx::Pie3DChart, :title => "fishery"
  end

  def teardown
  end

  def test_initialization
    assert_equal(@chart.view3D.rotX, 30, "view 3d default rotX incorrect")
    assert_equal(@chart.view3D.perspective, 30, "view 3d default perspective incorrect")
    assert_equal(@chart.series_type, Axlsx::PieSeries, "series type incorrect")
  end

  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    doc = Nokogiri::XML(@chart.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.empty?, "error free validation")
  end

end
