require 'tc_helper.rb'

class TestSeries < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    chart = @ws.add_chart Axlsx::Chart, :title => "fishery"
    @series = chart.add_series :title=>"bob"
  end

  def test_initialize
    assert_equal(@series.title.text, "bob", "series title has been applied")
    assert_equal(@series.order, @series.index, "order is index by default")
    assert_equal(@series.index, @series.chart.series.index(@series), "index is applied")
  end

  def test_order
    @series.order = 2
    assert_equal(@series.order, 2)
  end

end
