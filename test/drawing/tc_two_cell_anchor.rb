require 'test/unit'
require 'axlsx.rb'

class TestTwoCellAnchor < Test::Unit::TestCase
  def setup    
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @title = Axlsx::Title.new
    @chart = ws.add_chart Axlsx::Bar3DChart
    @anchor = @chart.graphic_frame.anchor
  end

  def teardown
  end

  def test_initialization
    assert(@anchor.from.col == 0)
    assert(@anchor.from.row == 0)
    assert(@anchor.to.col == 5)
    assert(@anchor.to.row == 10)
  end
  
  def test_start_at
    @anchor.start_at 5, 10
    assert(@anchor.from.col == 5)
    assert(@anchor.from.row == 10)
  end
  
  def test_end_at
    @anchor.end_at 10, 15
    assert(@anchor.to.col == 10)
    assert(@anchor.to.row == 15)
  end

  
end
