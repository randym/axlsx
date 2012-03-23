require 'tc_helper.rb'

class TestSeriesTitle < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @title = Axlsx::SeriesTitle.new
    @chart = ws.add_chart Axlsx::Bar3DChart
  end

  def teardown
  end

  def test_initialization
    assert(@title.text == "")
    assert(@title.cell == nil)
  end

  def test_text
    assert_raise(ArgumentError, "text must be a string") { @title.text = 123 }
    @title.cell = @row.cells.first
    @title.text = "bob"
    assert(@title.cell == nil, "setting title with text clears the cell")
  end

  def test_cell
    assert_raise(ArgumentError, "cell must be a Cell") { @title.cell = "123" }
    @title.cell = @row.cells.first
    assert(@title.text == "one")
  end

end
