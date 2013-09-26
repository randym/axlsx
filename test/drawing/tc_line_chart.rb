require 'tc_helper.rb'

class TestLineChart < Test::Unit::TestCase

  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @chart = ws.add_chart Axlsx::LineChart, :title => "fishery"
  end

  def teardown
  end

  def test_initialization
    assert_equal(@chart.grouping, :standard, "grouping defualt incorrect")
    assert_equal(@chart.series_type, Axlsx::LineSeries, "series type incorrect")
    assert(@chart.cat_axis.is_a?(Axlsx::CatAxis), "category axis not created")
    assert(@chart.val_axis.is_a?(Axlsx::ValAxis), "value access not created")
  end

  def test_grouping
    assert_raise(ArgumentError, "require valid grouping") { @chart.grouping = :inverted }
    assert_nothing_raised("allow valid grouping") { @chart.grouping = :stacked }
    assert(@chart.grouping == :stacked)
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
