require 'tc_helper.rb'

class TestAxis < Test::Unit::TestCase
  def setup
    @axis = Axlsx::Axis.new :gridlines => false, :title => 'Foo'
  end


  def test_initialization
    assert_equal(@axis.ax_pos, :b, "axis position default incorrect")
    assert_equal(@axis.tick_lbl_pos, :nextTo, "tick label position default incorrect")
    assert_equal(@axis.tick_lbl_pos, :nextTo, "tick label position default incorrect")
    assert_equal(@axis.crosses, :autoZero, "tick label position default incorrect")
    assert(@axis.scaling.is_a?(Axlsx::Scaling) && @axis.scaling.orientation == :minMax, "scaling default incorrect")
    assert_equal('Foo', @axis.title.text)
  end

  def test_color
    @axis.color = "00FF00"
    @axis.cross_axis = Axlsx::CatAxis.new
    str  = '<?xml version="1.0" encoding="UTF-8"?>'
    str << '<c:chartSpace xmlns:c="' << Axlsx::XML_NS_C << '" xmlns:a="' << Axlsx::XML_NS_A << '">'
    doc = Nokogiri::XML(@axis.to_xml_string(str)) 
    assert(doc.xpath("//a:srgbClr[@val='00FF00']"))
  end

  def test_cell_based_axis_title
    p = Axlsx::Package.new
    p.workbook.add_worksheet(:name=>'foosheet') do |sheet|
      sheet.add_row ['battle victories']
      sheet.add_row ['bird', 1, 2, 1]
      sheet.add_row ['cat', 7, 9, 10]
      sheet.add_chart(Axlsx::Line3DChart) do |chart|
        chart.add_series :data => sheet['B2:D2'], :labels => sheet['B1']
        chart.val_axis.title = sheet['A1']
        assert_equal('battle victories', chart.val_axis.title.text)
      end
    end
  end

  def test_axis_position
    assert_raise(ArgumentError, "requires valid axis position") { @axis.ax_pos = :nowhere }
    assert_nothing_raised("accepts valid axis position") { @axis.ax_pos = :r }
  end

  def test_label_rotation
    assert_raise(ArgumentError, "requires valid angle") { @axis.label_rotation = :nowhere }
    assert_raise(ArgumentError, "requires valid angle") { @axis.label_rotation = 91 }
    assert_raise(ArgumentError, "requires valid angle") { @axis.label_rotation = -91 }
    assert_nothing_raised("accepts valid angle") { @axis.label_rotation = 45 }
    assert_equal(@axis.label_rotation, 45 * 60000)
  end

  def test_tick_label_position
    assert_raise(ArgumentError, "requires valid tick label position") { @axis.tick_lbl_pos = :nowhere }
    assert_nothing_raised("accepts valid tick label position") { @axis.tick_lbl_pos = :high }
  end

  def test_format_code
    assert_raise(ArgumentError, "requires valid format code") { @axis.format_code = :high }
    assert_nothing_raised("accepts valid format code") { @axis.format_code = "00.##"  }
  end

  def create_chart_with_formatting(format_string=nil)
    p = Axlsx::Package.new
    p.workbook.add_worksheet(:name => "Formatting Test") do |sheet|
      sheet.add_row(['test', 20])
      sheet.add_chart(Axlsx::Bar3DChart, :start_at => [0,5], :end_at => [10, 20], :title => "Test Formatting") do |chart|
        chart.add_series :data => sheet["B1:B1"], :labels => sheet["A1:A1"]
        chart.val_axis.format_code = format_string if format_string
        doc = Nokogiri::XML(chart.to_xml_string)
        yield doc
      end
    end
  end

  def test_format_code_resets_source_linked
    create_chart_with_formatting("#,##0.00") do |doc|
      assert_equal(doc.xpath("//c:valAx/c:numFmt[@formatCode='#,##0.00']").size, 1)
      assert_equal(doc.xpath("//c:valAx/c:numFmt[@sourceLinked='0']").size, 1)
    end
  end

  def test_no_format_code_keeps_source_linked
    create_chart_with_formatting do |doc|
      assert_equal(doc.xpath("//c:valAx/c:numFmt[@formatCode='General']").size, 1)
      assert_equal(doc.xpath("//c:valAx/c:numFmt[@sourceLinked='1']").size, 1)
    end
  end

  def test_crosses
    assert_raise(ArgumentError, "requires valid crosses") { @axis.crosses = 1 }
    assert_nothing_raised("accepts valid crosses") { @axis.crosses = :min }
  end

  def test_gridlines
    assert_raise(ArgumentError, "requires valid gridlines") { @axis.gridlines = 'alice' }
    assert_nothing_raised("accepts valid crosses") { @axis.gridlines = false }
  end
  
  def test_to_xml_string
    @axis.cross_axis = Axlsx::CatAxis.new
    str  = '<?xml version="1.0" encoding="UTF-8"?>'
    str << '<c:chartSpace xmlns:c="' << Axlsx::XML_NS_C << '" xmlns:a="' << Axlsx::XML_NS_A << '">'
    doc = Nokogiri::XML(@axis.to_xml_string(str)) 
    assert(doc.xpath('//a:noFill'))
    assert(doc.xpath("//c:crosses[@val='#{@axis.crosses.to_s}']"))
    assert(doc.xpath("//c:crossAx[@val='#{@axis.cross_axis.to_s}']"))
    assert(doc.xpath("//a:bodyPr[@rot='#{@axis.label_rotation.to_s}']"))
    assert(doc.xpath("//a:t[text()='Foo']"))
  end
end
