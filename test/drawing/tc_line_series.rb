require 'test/unit'
require 'axlsx.rb'

class TestLineSeries < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    chart = @ws.drawing.add_chart Axlsx::Line3DChart, :title => "fishery"
    @series = chart.add_series :data=>[0,1,2], :labels=>["zero", "one", "two"], :title=>"bob"
  end
  
  def test_initialize
    assert_equal(@series.title.text, "bob", "series title has been applied")
    assert_equal(@series.data, [0,1,2], "data option applied")
    assert_equal(@series.labels, ["zero", "one","two"], "labels option applied")    
  end

  def test_data
    assert_equal(@series.data, [0,1,2])
  end

  def test_labels
    assert_equal(@series.labels, ["zero", "one", "two"])
  end

end
