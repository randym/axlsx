require 'test/unit'
require 'axlsx.rb'

class TestRow < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    @row = @ws.add_row
  end
  
  def test_initialize
    assert(@row.cells.empty?, "no cells by default")
    assert_equal(@row.worksheet, @ws, "has a reference to the worksheet")
  end

  def test_index
    assert_equal(@row.index, @row.worksheet.rows.index(@row))
  end

  def test_add_cell
    c = @row.add_cell(1)
    assert_equal(@row.cells.last, c)
  end

  def test_array_to_cells
    r = @ws.add_row [1,2,3], :style=>0, :types=>:integer
    assert_equal(r.cells.size, 3)
  end
end
