require 'tc_helper.rb'

class TestCatAxisData < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet
    chart = @ws.drawing.add_chart Axlsx::Bar3DChart
    @series = chart.add_series :labels=>["zero", "one", "two"]
  end

  def test_initialize
    assert(@series.labels.is_a?Axlsx::SimpleTypedList)
    assert_equal(@series.labels, ["zero", "one", "two"])
  end

end
