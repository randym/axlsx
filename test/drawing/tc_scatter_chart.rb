require 'tc_helper.rb'

class TestScatterChart < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @chart = ws.add_chart Axlsx::ScatterChart, :title => "A Title"
  end

  def teardown
  end

  def test_initialization
    assert_equal(@chart.scatterStyle, :lineMarker, "scatterStyle defualt incorrect")
    assert_equal(@chart.series_type, Axlsx::ScatterSeries, "series type incorrect")
    assert(@chart.xValAxis.is_a?(Axlsx::ValAxis), "independant value axis not created")
    assert(@chart.yValAxis.is_a?(Axlsx::ValAxis), "dependant value axis not created")
  end

  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    doc = Nokogiri::XML(@chart.to_xml)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.empty?, "error free validation")
  end

end
