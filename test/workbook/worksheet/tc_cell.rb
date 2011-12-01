require 'test/unit'
require 'axlsx.rb'

class TestCell < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    p.workbook.styles.add_style :sz=>20
    @row = @ws.add_row
    @c = @row.add_cell 1, :type=>:float, :style=>1
  end
  
  def test_initialize
    assert_equal(@row.cells.last, @c, "the cell was added to the row")
    assert_equal(@c.type, :float, "type option is applied")
    assert_equal(@c.style, 1, "style option is applied")
    assert_equal(@c.value, 1.0, "type option is applied and value is casted")
  end

  def test_style_date_data
    c = Axlsx::Cell.new(@c.row, Time.now)
    assert_equal(Axlsx::STYLE_DATE, c.style)
  end

  def test_index
    assert_equal(@c.index, @row.cells.index(@c))
  end

  def test_r
    assert_equal(@c.r, "A1", "calculate cell reference")
  end

  def test_r_abs
    assert_equal(@c.r_abs,"$A$1", "calculate absolute cell reference")
  end

  def test_style
    assert_raise(ArgumentError, "must reject invalid style indexes") { @c.style=@c.row.worksheet.workbook.styles.cellXfs.size }
    assert_nothing_raised("must allow valid style index changes") {@c.style=1} 
    assert_equal(@c.style, 1)
  end

  def test_type
    assert_raise(ArgumentError, "type must be :string, :integer, :float, :time") { @c.type = :array }
    assert_nothing_raised("type can be changed") { @c.type = :string }
    assert_equal(@c.value, "1.0", "changing type casts the value")
    
    assert_equal(@row.add_cell(Time.now).type, :time, 'time should be time')
  end

  def test_value
    assert_raise(ArgumentError, "type must be :string, :integer, :float, :time") { @c.type = :array }
    assert_nothing_raised("type can be changed") { @c.type = :string }
    assert_equal(@c.value, "1.0", "changing type casts the value")
  end

  def test_col_ref
    assert_equal(@c.send(:col_ref), "A")
  end

  def test_cell_type_from_value
    assert_equal(@c.send(:cell_type_from_value, 1.0), :float)
    assert_equal(@c.send(:cell_type_from_value, 1), :integer)
    assert_equal(@c.send(:cell_type_from_value, Time.now), :time)
    assert_equal(@c.send(:cell_type_from_value, []), :string)
    assert_equal(@c.send(:cell_type_from_value, "d"), :string)
    assert_equal(@c.send(:cell_type_from_value, nil), :string)
    assert_equal(@c.send(:cell_type_from_value, -1), :integer)
  end

  def test_cast_value    
    @c.type = :string
    assert_equal(@c.send(:cast_value, 1.0), "1.0")
    @c.type = :integer
    assert_equal(@c.send(:cast_value, 1.0), 1)
    @c.type = :float
    assert_equal(@c.send(:cast_value, "1.0"), 1.0)
    @c.type = :string
    assert_equal(@c.send(:cast_value, nil), "")

  end
end
