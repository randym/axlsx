require 'tc_helper.rb'

class TestBubbleSeries < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    @chart = @ws.add_chart Axlsx::BubbleChart, :title => "Bubble Chart"
    @series = @chart.add_series :xData=>[1,2,4], :yData=>[1,3,9], :bubbleSize=>[1,5,7], :title=>"GDP", :color => 'FF0000'
  end

  def test_initialize
    assert_equal(@series.title.text, "GDP", "series title has been applied")
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@chart.to_xml_string)
    assert_equal(doc.xpath("//a:srgbClr[@val='#{@series.color}']").size,2)
  end

end
