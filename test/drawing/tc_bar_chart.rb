require 'tc_helper.rb'
require 'axlsx/drawing/bar_chart'

class TestBarChart < Test::Unit::TestCase

  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.new]
    @chart = ws.add_chart Axlsx::BarChart, :title => "fishery", :bar_dir => :col
  end

  def test_initialization
    assert_equal(@chart.series_type, Axlsx::BarSeries, "series type incorrect")
  end

  def test_chart_type
    doc = Nokogiri::XML(@chart.to_xml_string)
    assert_equal(1, doc.xpath('//c:barChart').size)
  end

  def test_bar_direction
    assert_raise(ArgumentError, "require valid bar direction") { @chart.bar_dir = :left }
    assert_nothing_raised("allow valid bar direction") { @chart.bar_dir = :col }
    assert(@chart.bar_dir == :col)
  end

  def test_grouping
    assert_raise(ArgumentError, "require valid grouping") { @chart.grouping = :inverted }
    assert_nothing_raised("allow valid grouping") { @chart.grouping = :standard }
    assert(@chart.grouping == :standard)
  end

  def test_gapWidth
    assert_raise(ArgumentError, "require valid gap width") { @chart.gap_width = 200 }
    assert_nothing_raised("allow valid gapWidth") { @chart.gap_width = "200%" }
    assert(@chart.gap_width == "200%")
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

  def test_to_xml_string_has_axes_in_correct_order
    str = @chart.to_xml_string
    cat_axis_position = str.index(@chart.axes[:cat_axis].id.to_s)
    val_axis_position = str.index(@chart.axes[:val_axis].id.to_s)
    assert(cat_axis_position < val_axis_position, "cat_axis must occur earlier than val_axis in the XML")
  end

end