require 'tc_helper.rb'

class TestLineSeries < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    chart = @ws.add_chart Axlsx::Line3DChart, :title => "fishery"
    @series = chart.add_series :data=>[0,1,2], :labels=>["zero", "one", "two"], :title=>"bob", :color => "#FF0000", :show_marker => true
  end

  def test_initialize
    assert_equal(@series.title.text, "bob", "series title has been applied")
    assert_equal(@series.labels.class, Axlsx::AxDataSource)
    assert_equal(@series.data.class, Axlsx::NumDataSource)

  end

  def test_show_marker
    assert_equal(true, @series.show_marker)
    @series.show_marker = false
    assert_equal(false, @series.show_marker)
  end  
  def test_to_xml_string
    doc = Nokogiri::XML(@series.to_xml_string)
    assert(doc.xpath("//srgbClr[@val='#{@series.color}']"))
    assert(doc.xpath("//marker"))
  end
  #TODO serialization testing
end
