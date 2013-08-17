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
    @chart.title = @row.cells.first
    assert_equal(@chart.title.text, "one", "the title text was set via cell reference")
    assert_equal(@chart.title.cell, @row.cells.first)
  end

  def test_to_from_marker_access
    assert(@chart.to.is_a?(Axlsx::Marker))
    assert(@chart.from.is_a?(Axlsx::Marker))
  end

  def test_style
    assert_raise(ArgumentError) { @chart.style = 49 }
    assert_nothing_raised { @chart.style = 2 }
    assert_equal(@chart.style, 2)
  end
  
  def test_vary_colors
    assert_equal(true, @chart.vary_colors)
    assert_raise(ArgumentError) { @chart.vary_colors = 7 }
    assert_nothing_raised { @chart.vary_colors = false }
    assert_equal(false, @chart.vary_colors)
  end

  def test_display_blanks_as
    assert_equal(:gap, @chart.display_blanks_as, "default is not :gap")
    assert_raise(ArgumentError, "did not validate possible values") { @chart.display_blanks_as = :hole }
    assert_nothing_raised { @chart.display_blanks_as = :zero }
    assert_nothing_raised { @chart.display_blanks_as = :span }
    assert_equal(:span, @chart.display_blanks_as)
  end

  def test_start_at
    @chart.start_at 15, 25
    assert_equal(@chart.graphic_frame.anchor.from.col, 15)
    assert_equal(@chart.graphic_frame.anchor.from.row, 25)
    @chart.start_at @row.cells.first
    assert_equal(@chart.graphic_frame.anchor.from.col, 0)
    assert_equal(@chart.graphic_frame.anchor.from.row, 0)
    @chart.start_at [5,6]
    assert_equal(@chart.graphic_frame.anchor.from.col, 5)
    assert_equal(@chart.graphic_frame.anchor.from.row, 6)
    
  end

  def test_end_at
    @chart.end_at 25, 90
    assert_equal(@chart.graphic_frame.anchor.to.col, 25)
    assert_equal(@chart.graphic_frame.anchor.to.row, 90)
    @chart.end_at @row.cells.last
    assert_equal(@chart.graphic_frame.anchor.to.col, 2)
    assert_equal(@chart.graphic_frame.anchor.to.row, 0)
    @chart.end_at [10,11]
    assert_equal(@chart.graphic_frame.anchor.to.col, 10)
    assert_equal(@chart.graphic_frame.anchor.to.row, 11)
  
  end

  def test_add_series
    s = @chart.add_series :data=>[0,1,2,3], :labels => ["one", 1, "anything"], :title=>"bob"
    assert_equal(@chart.series.last, s, "series has been added to chart series collection")
    assert_equal(s.title.text, "bob", "series title has been applied")
  end

  def test_pn
    assert_equal(@chart.pn, "charts/chart1.xml")
  end
  
  def test_d_lbls
    assert_equal(nil, @chart.instance_values[:d_lbls])
    @chart.d_lbls.d_lbl_pos = :t
    assert(@chart.d_lbls.is_a?(Axlsx::DLbls), 'DLbls instantiated on access')
  end
  
  def test_to_xml_string
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    doc = Nokogiri::XML(@chart.to_xml_string)
    errors = schema.validate(doc).map { |error| puts error.message; error }
    assert(errors.empty?, "error free validation")
  end

  def test_to_xml_string_for_display_blanks_as
    @chart.display_blanks_as = :span
    doc = Nokogiri::XML(@chart.to_xml_string)
    assert_equal("span", doc.xpath("//c:dispBlanksAs").attr("val").value, "did not use the display_blanks_as configuration")
  end
end
