require 'test/unit'
require 'axlsx.rb'

class TestBarSeries < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    chart = @ws.drawing.add_chart Axlsx::Bar3DChart, :title => "fishery"
    @series = chart.add_series :data=>[0,1,2], :labels=>["zero", "one", "two"], :title=>"bob"
  end
  
  def test_initialize
    assert_equal(@series.title, "bob", "series title has been applied")
    assert_equal(@series.data, [0,1,2], "data option applied")
    assert_equal(@series.labels, ["zero", "one","two"], "labels option applied")    
    assert_equal(@series.shape, :box, "series shape has been applied")   
  end

  def test_data
    assert_equal(@series.data, [0,1,2])
  end

  def test_labels
    assert_equal(@series.labels, ["zero", "one", "two"])
  end

  def test_shape
    assert_raise(ArgumentError, "require valid shape") { @series.shape = :teardropt }
    assert_nothing_raised("allow valid shape") { @series.shape = :cone }
    assert(@series.shape == :cone)
  end

end
