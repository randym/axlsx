require 'tc_helper.rb'

class TestBubbleChart < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    @chart = nil
    @p.workbook.add_worksheet do |sheet|
      sheet.add_row ["First",  1,  5,  7,  9]
      sheet.add_row ["",       1, 25, 49, 81]
      sheet.add_row ["",       1, 42, 60, 75]
      sheet.add_row ["Second", 5,  2, 14,  9]
      sheet.add_row ["",       5, 10, 15, 20]
      sheet.add_row ["",       5, 28, 92, 13]
      sheet.add_chart(Axlsx::BubbleChart, :title => "example: Bubble Chart") do |chart|
        chart.start_at 0, 4
        chart.end_at 10, 19
        chart.add_series :xData => sheet["B1:E1"], :yData => sheet["B2:E2"], :bubbleSize => sheet["B3:E3"], :title => sheet["A1"]
        chart.add_series :xData => sheet["B4:E4"], :yData => sheet["B5:E5"], :bubbleSize => sheet["B6:E6"], :title => sheet["A3"]
        @chart = chart
      end
    end
  end

  def teardown
  end

  def test_initialization
    assert_equal(@chart.series_type, Axlsx::BubbleSeries, "series type incorrect")
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
