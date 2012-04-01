require 'tc_helper.rb'

class TestBar3DChart < Test::Unit::TestCase

  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @chart = ws.add_chart Axlsx::Bar3DChart, :title => "fishery"
  end

  def teardown
  end

  def test_initialization
    assert_equal(@chart.grouping, :clustered, "grouping defualt incorrect")
    assert_equal(@chart.series_type, Axlsx::BarSeries, "series type incorrect")
    assert_equal(@chart.barDir, :bar, " bar direction incorrect")
    assert(@chart.catAxis.is_a?(Axlsx::CatAxis), "category axis not created")
    assert(@chart.valAxis.is_a?(Axlsx::ValAxis), "value access not created")
  end

  def test_bar_direction
    assert_raise(ArgumentError, "require valid bar direction") { @chart.barDir = :left }
    assert_nothing_raised("allow valid bar direction") { @chart.barDir = :col }
    assert(@chart.barDir == :col)
  end

 def test_grouping
   assert_raise(ArgumentError, "require valid grouping") { @chart.grouping = :inverted }
   assert_nothing_raised("allow valid grouping") { @chart.grouping = :standard }
   assert(@chart.grouping == :standard)
 end


 def test_gapWidth
   assert_raise(ArgumentError, "require valid gap width") { @chart.gapWidth = 200 }
   assert_nothing_raised("allow valid gapWidth") { @chart.gapWidth = "200%" }
   assert(@chart.gapWidth == "200%")
 end

 def test_gapDepth
   assert_raise(ArgumentError, "require valid gapDepth") { @chart.gapDepth = 200 }
   assert_nothing_raised("allow valid gapDepth") { @chart.gapDepth = "200%" }
   assert(@chart.gapDepth == "200%")
 end

  def test_shape
    assert_raise(ArgumentError, "require valid shape") { @chart.shape = :star }
    assert_nothing_raised("allow valid shape") { @chart.shape = :cone }
    assert(@chart.shape == :cone)
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
