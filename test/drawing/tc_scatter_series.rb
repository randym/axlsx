require 'test/unit'
require 'axlsx.rb'

class TestScatterSeries < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    chart = @ws.drawing.add_chart Axlsx::ScatterChart, :title => "Scatter Chart"
    @series = chart.add_series :xData=>[1,2,4], :yData=>[1,3,9], :title=>"exponents"
  end

  def test_initialize
    assert_equal(@series.title.text, "exponents", "series title has been applied")
  end

  def test_data
    assert_equal(@series.xData, [1,2,4])
    assert_equal(@series.yData, [1,3,9])
  end
end
