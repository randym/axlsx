require 'tc_helper.rb'

class TestCell < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    p.workbook.styles.add_style :sz=>20
    @row = @ws.add_row
    @c = @row.add_cell 1, :type=>:float, :style=>1
    data = (0..26).map { |index| index }
    @ws.add_row data
    @cAA = @ws["AA2"]
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

  def test_row
    assert_equal(@c.row, @row)
  end

  def test_index
    assert_equal(@c.index, @row.cells.index(@c))
  end

  def test_index
    assert_equal(@c.pos, [@c.index, @c.row.index])
  end

  def test_r
    assert_equal(@c.r, "A1", "calculate cell reference")
  end

  def test_wide_r
      assert_equal(@cAA.r, "AA2", "calculate cell reference")
  end

  def test_r_abs
    assert_equal(@c.r_abs,"$A$1", "calculate absolute cell reference")
    assert_equal(@cAA.r_abs,"$AA$2", "needs to accept multi-digit columns")
  end

  def test_style
    assert_raise(ArgumentError, "must reject invalid style indexes") { @c.style=@c.row.worksheet.workbook.styles.cellXfs.size }
    assert_nothing_raised("must allow valid style index changes") {@c.style=1}
    assert_equal(@c.style, 1)
  end

  def test_type
    assert_raise(ArgumentError, "type must be :string, :integer, :float, :date, :time, :boolean") { @c.type = :array }
    assert_nothing_raised("type can be changed") { @c.type = :string }
    assert_equal(@c.value, "1.0", "changing type casts the value")

    assert_equal(@row.add_cell(Time.now).type, :time, 'time should be time')
    assert_equal(@row.add_cell(Date.today).type, :date, 'date should be date')
    assert_equal(@row.add_cell(true).type, :boolean, 'boolean should be boolean')
  end

  def test_value
    assert_raise(ArgumentError, "type must be :string, :integer, :float, :date, :time, :boolean") { @c.type = :array }
    assert_nothing_raised("type can be changed") { @c.type = :string }
    assert_equal(@c.value, "1.0", "changing type casts the value")
  end

  def test_col_ref
    #TODO move to axlsx spec
    assert_equal(Axlsx.col_ref(0), "A")
  end

  def test_cell_type_from_value
    assert_equal(@c.send(:cell_type_from_value, 1.0), :float)
    assert_equal(@c.send(:cell_type_from_value, 1), :integer)
    assert_equal(@c.send(:cell_type_from_value, Date.today), :date)
    assert_equal(@c.send(:cell_type_from_value, Time.now), :time)
    assert_equal(@c.send(:cell_type_from_value, []), :string)
    assert_equal(@c.send(:cell_type_from_value, "d"), :string)
    assert_equal(@c.send(:cell_type_from_value, nil), :string)
    assert_equal(@c.send(:cell_type_from_value, -1), :integer)
    assert_equal(@c.send(:cell_type_from_value, true), :boolean)
    assert_equal(@c.send(:cell_type_from_value, false), :boolean)
  end

  def test_cast_value
    @c.type = :string
    assert_equal(@c.send(:cast_value, 1.0), "1.0")
    @c.type = :integer
    assert_equal(@c.send(:cast_value, 1.0), 1)
    @c.type = :float
    assert_equal(@c.send(:cast_value, "1.0"), 1.0)
    @c.type = :string
    assert_equal(@c.send(:cast_value, nil), nil)
    @c.type = :float
    assert_equal(@c.send(:cast_value, nil), nil)
    @c.type = :boolean
    assert_equal(@c.send(:cast_value, true), 1)
    assert_equal(@c.send(:cast_value, false), 0)
  end

  def test_color
    assert_raise(ArgumentError) { @c.color = -1.1 }
    assert_nothing_raised { @c.color = "FF00FF00" }
    assert_equal(@c.color.rgb, "FF00FF00")
  end

  def test_scheme
    assert_raise(ArgumentError) { @c.scheme = -1.1 }
    assert_nothing_raised { @c.scheme = :major }
    assert_equal(@c.scheme, :major)
  end

  def test_vertAlign
    assert_raise(ArgumentError) { @c.vertAlign = -1.1 }
    assert_nothing_raised { @c.vertAlign = :baseline }
    assert_equal(@c.vertAlign, :baseline)
  end

  def test_sz
    assert_raise(ArgumentError) { @c.sz = -1.1 }
    assert_nothing_raised { @c.sz = 12 }
    assert_equal(@c.sz, 12)
  end

  def test_extend
    assert_raise(ArgumentError) { @c.extend = -1.1 }
    assert_nothing_raised { @c.extend = false }
    assert_equal(@c.extend, false)
  end

  def test_condense
    assert_raise(ArgumentError) { @c.condense = -1.1 }
    assert_nothing_raised { @c.condense = false }
    assert_equal(@c.condense, false)
  end

  def test_shadow
    assert_raise(ArgumentError) { @c.shadow = -1.1 }
    assert_nothing_raised { @c.shadow = false }
    assert_equal(@c.shadow, false)
  end

  def test_outline
    assert_raise(ArgumentError) { @c.outline = -1.1 }
    assert_nothing_raised { @c.outline = false }
    assert_equal(@c.outline, false)
  end

  def test_strike
    assert_raise(ArgumentError) { @c.strike = -1.1 }
    assert_nothing_raised { @c.strike = false }
    assert_equal(@c.strike, false)
  end

  def test_u
    assert_raise(ArgumentError) { @c.u = -1.1 }
    assert_nothing_raised { @c.u = false }
    assert_equal(@c.u, false)
  end

  def test_i
    assert_raise(ArgumentError) { @c.i = -1.1 }
    assert_nothing_raised { @c.i = false }
    assert_equal(@c.i, false)
  end

  def test_rFont
    assert_raise(ArgumentError) { @c.font_name = -1.1 }
    assert_nothing_raised { @c.font_name = "Arial" }
    assert_equal(@c.font_name, "Arial")
  end

  def test_charset
    assert_raise(ArgumentError) { @c.charset = -1.1 }
    assert_nothing_raised { @c.charset = 1 }
    assert_equal(@c.charset, 1)
  end

  def test_family
    assert_raise(ArgumentError) { @c.family = -1.1 }
    assert_nothing_raised { @c.family = "Who knows!" }
    assert_equal(@c.family, "Who knows!")
  end

  def test_b
    assert_raise(ArgumentError) { @c.b = -1.1 }
    assert_nothing_raised { @c.b = false }
    assert_equal(@c.b, false)
  end

  def test_merge_with_string
    assert_equal(@c.row.worksheet.merged_cells.size, 0)
    @c.row.add_cell 2
    @c.row.add_cell 3
    @c.merge "A2"
    assert_equal(@c.row.worksheet.merged_cells.last, "A1:A2")
  end

  def test_merge_with_cell
    assert_equal(@c.row.worksheet.merged_cells.size, 0)
    @c.row.add_cell 2
    @c.row.add_cell 3
    @c.merge @row.cells.last
    assert_equal(@c.row.worksheet.merged_cells.last, "A1:C1")
  end

  def test_ssti
    assert_raise(ArgumentError, "ssti must be an unsigned integer!") { @c.send(:ssti=, -1) }
    @c.send :ssti=, 1
    assert_equal(@c.ssti, 1)
  end

  def test_plain_string
    @c.type = :integer
    assert_equal(@c.plain_string?, false)

    @c.type = :string
    @c.value = 'plain string'
    assert_equal(@c.plain_string?, true)

    @c.value = nil
    assert_equal(@c.plain_string?, false)

    @c.value = ''
    assert_equal(@c.plain_string?, false)

    @c.value = '=sum'
    assert_equal(@c.plain_string?, false)

    @c.value = 'plain string'
    @c.font_name = 'Arial'
    assert_equal(@c.plain_string?, false)
  end

  def test_to_xml_string
    c_xml = Nokogiri::XML(@c.to_xml_string(1,1))
    assert_equal(c_xml.xpath("/c[@s=1]").size, 1)
  end
  def test_to_xml
    # TODO This could use some much more stringent testing related to the xml content generated!
    row = @ws.add_row [Time.now, Date.today, true, 1, 1.0, "text", "=sum(A1:A2)"]
    schema = Nokogiri::XML::Schema(File.open(Axlsx::SML_XSD))
    doc = Nokogiri::XML(@ws.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end
    assert(errors.empty?, "error free validation")

  end

end
