require 'test/unit'
require 'axlsx.rb'

class TestValAxisData < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
    chart = @ws.drawing.add_chart Axlsx::Line3DChart
    @series = chart.add_series :data=>[0,1,2]
  end
  
  def test_initialize
    assert(@series.data.is_a?Axlsx::SimpleTypedList)
    assert_equal(@series.data, [0,1,2])
  end

end
