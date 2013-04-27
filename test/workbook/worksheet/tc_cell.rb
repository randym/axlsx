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

  def test_pos
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
    assert_equal(:float, @row.add_cell(1.0/10**7).type, 'properly identify exponential floats as float type')
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
    assert_equal(@c.send(:cell_type_from_value, 1.0/10**6), :float)
    assert_equal(:iso_8601, @c.send(:cell_type_from_value, '2008-08-30T01:45:36.123+09:00'))
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
    @c.type = :iso_8601
    assert_equal("2012-10-10T12:24", @c.send(:cast_value, "2012-10-10T12:24"))
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
    @c.type = :string
    assert_raise(ArgumentError) { @c.u = -1.1 }
    assert_nothing_raised { @c.u = :single }
    assert_equal(@c.u, :single)
    doc = Nokogiri::XML(@c.to_xml_string(1,1))
    assert(doc.xpath('//u[@val="single"]'))
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
    assert_nothing_raised { @c.family = 5 }
    assert_equal(@c.family, 5)
  end

  def test_b
    assert_raise(ArgumentError) { @c.b = -1.1 }
    assert_nothing_raised { @c.b = false }
    assert_equal(@c.b, false)
  end

  def test_merge_with_string
    @c.row.add_cell 2
    @c.row.add_cell 3
    @c.merge "A2"
    assert_equal(@c.row.worksheet.send(:merged_cells).last, "A1:A2")
  end

  def test_merge_with_cell
    @c.row.add_cell 2
    @c.row.add_cell 3
    @c.merge @row.cells.last
    assert_equal(@c.row.worksheet.send(:merged_cells).last, "A1:C1")
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

  def test_to_xml_string_nil
    @c.value = nil
    c_xml = Nokogiri::XML(@c.to_xml_string(1,1))
    assert_equal(c_xml.xpath("/c[@s=1]").size, 1)
  end
  
  def test_to_xml_string_with_run
    # Actually quite a number of similar run styles
    # but the processing should be the same
    @c.b = true
    @c.type = :string
    @c.value = "a"
    @c.font_name = 'arial'
    @c.color = 'FF0000'
    c_xml = Nokogiri::XML(@c.to_xml_string(1,1))
    assert(c_xml.xpath("//b"))
  end

  def test_to_xml_string_formula
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet do |sheet|
      sheet.add_row ["=IF(2+2=4,4,5)"]
    end
    doc = Nokogiri::XML(ws.to_xml_string)
    assert(doc.xpath("//f[@text()='IF(2+2=4,4,5)']"))

  end

  def test_font_size_with_custom_style_and_no_sz
    @c.style = @c.row.worksheet.workbook.styles.add_style :bg_color => 'FF00FF'
    sz = @c.send(:font_size)
    assert_equal(sz, @c.row.worksheet.workbook.styles.fonts.first.sz)
  end

  def test_font_size_with_bolding
    @c.style = @c.row.worksheet.workbook.styles.add_style :b => true
    assert_equal(@c.row.worksheet.workbook.styles.fonts.first.sz * 1.5, @c.send(:font_size))
  end

  def test_font_size_with_custom_sz
    @c.style = @c.row.worksheet.workbook.styles.add_style :sz => 52
    sz = @c.send(:font_size)
    assert_equal(sz, 52)
  end
  

  def test_cell_with_sz
    @c.sz = 25
    assert_equal(25, @c.send(:font_size))
  end
  def test_to_xml
    # TODO This could use some much more stringent testing related to the xml content generated!
    @ws.add_row [Time.now, Date.today, true, 1, 1.0, "text", "=sum(A1:A2)", "2013-01-13T13:31:25.123"]
    @ws.rows.last.cells[5].u = true
    
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
