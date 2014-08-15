require 'tc_helper.rb'

class TestScatterSeries < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    @chart = @ws.add_chart Axlsx::ScatterChart, :title => "Scatter Chart"
    @series = @chart.add_series :xData=>[1,2,4], :yData=>[1,3,9], :title=>"exponents", :color => 'FF0000', :smooth => true
  end

  def test_initialize
    assert_equal(@series.title.text, "exponents", "series title has been applied")
  end

  def test_smoothed_chart_default_smoothing
    @chart = @ws.add_chart Axlsx::ScatterChart, :title => "Smooth Chart", :scatter_style => :smoothMarker
    @series = @chart.add_series :xData=>[1,2,4], :yData=>[1,3,9], :title=>"smoothed exponents"
    assert(@series.smooth, "series is smooth by default on smooth charts")
  end
  
  def test_unsmoothed_chart_default_smoothing
    @chart = @ws.add_chart Axlsx::ScatterChart, :title => "Unsmooth Chart", :scatter_style => :line
    @series = @chart.add_series :xData=>[1,2,4], :yData=>[1,3,9], :title=>"unsmoothed exponents"
    assert(!@series.smooth, "series is not smooth by default on non-smooth charts")
  end
  
  def test_explicit_smoothing
    @chart = @ws.add_chart Axlsx::ScatterChart, :title => "Unsmooth Chart, Smooth Series", :scatter_style => :line
    @series = @chart.add_series :xData=>[1,2,4], :yData=>[1,3,9], :title=>"smoothed exponents", :smooth => true
    assert(@series.smooth, "series is smooth when overriding chart default")
  end    
  
  def test_explicit_unsmoothing
    @chart = @ws.add_chart Axlsx::ScatterChart, :title => "Smooth Chart, Unsmooth Series", :scatter_style => :smoothMarker
    @series = @chart.add_series :xData=>[1,2,4], :yData=>[1,3,9], :title=>"unsmoothed exponents", :smooth => false
    assert(!@series.smooth, "series is not smooth when overriding chart default")
  end    

  def test_to_xml_string
    doc = Nokogiri::XML(@chart.to_xml_string)
    assert_equal(doc.xpath("//a:srgbClr[@val='#{@series.color}']").size,4)
  end

end
