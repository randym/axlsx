require 'tc_helper.rb'

class TestScatterChart < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    @chart = nil
    @p.workbook.add_worksheet do |sheet|
      sheet.add_row ["First",  1,  5,  7,  9]
      sheet.add_row ["",       1, 25, 49, 81]
      sheet.add_row ["Second", 5,  2, 14,  9]
      sheet.add_row ["",       5, 10, 15, 20]
      sheet.add_chart(Axlsx::ScatterChart, :title => "example 7: Scatter Chart") do |chart|
        chart.start_at 0, 4
        chart.end_at 10, 19
        chart.add_series :xData => sheet["B1:E1"], :yData => sheet["B2:E2"], :title => sheet["A1"]
        chart.add_series :xData => sheet["B3:E3"], :yData => sheet["B4:E4"], :title => sheet["A3"]
        @chart = chart
      end
    end
  end

  def teardown
  end

  def test_scatter_style
    @chart.scatterStyle = :marker
    assert(@chart.scatterStyle == :marker)
    assert_raise(ArgumentError) { @chart.scatterStyle = :buckshot }
  end
  def test_initialization
    assert_equal(@chart.scatterStyle, :lineMarker, "scatterStyle defualt incorrect")
    assert_equal(@chart.series_type, Axlsx::ScatterSeries, "series type incorrect")
    assert(@chart.xValAxis.is_a?(Axlsx::ValAxis), "independant value axis not created")
    assert(@chart.yValAxis.is_a?(Axlsx::ValAxis), "dependant value axis not created")
  end

  def test_to_xml_string
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
