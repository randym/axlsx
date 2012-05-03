require 'tc_helper.rb'

class TestChart < Test::Unit::TestCase

  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @chart = ws.add_chart Axlsx::Bar3DChart, :title => "fishery"
  end

  def teardown
  end

  def test_initialization
    assert_equal(@p.workbook.charts.last,@chart, "the chart is in the workbook")
    assert_equal(@chart.title.text, "fishery", "the title option has been applied")
    assert((@chart.series.is_a?(Axlsx::SimpleTypedList) && @chart.series.empty?), "The series is initialized and empty")
  end

  def test_title
    @chart.title.text = 'wowzer'
    assert_equal(@chart.title.text, "wowzer", "the title text via a string")
    assert_equal(@chart.title.cell, nil, "the title cell is nil as we set the title with text.")
    @chart.title.cell = @row.cells.first
    assert_equal(@chart.title.text, "one", "the title text was set via cell reference")
    assert_equal(@chart.title.cell, @row.cells.first)
  end

  def test_style
    assert_raise(ArgumentError) { @chart.style = 49 }
    assert_nothing_raised { @chart.style = 2 }
    assert_equal(@chart.style, 2)
  end

  def test_start_at
    @chart.start_at 15,25
    assert_equal(@chart.graphic_frame.anchor.from.col, 15)
    assert_equal(@chart.graphic_frame.anchor.from.row, 25)

  end

  def test_end_at
    @chart.end_at 25, 90
    assert_equal(@chart.graphic_frame.anchor.to.col, 25)
    assert_equal(@chart.graphic_frame.anchor.to.row, 90)
  end

  def test_add_series
    s = @chart.add_series :data=>[0,1,2,3], :labels => ["one", 1, "anything"], :title=>"bob"
    assert_equal(@chart.series.last, s, "series has been added to chart series collection")
    assert_equal(s.title.text, "bob", "series title has been applied")
  end

  def test_pn
    assert_equal(@chart.pn, "charts/chart1.xml")
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
