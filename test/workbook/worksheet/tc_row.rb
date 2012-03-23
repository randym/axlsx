require 'tc_helper.rb'

class TestRow < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name=>"hmmm"
    @row = @ws.add_row
  end

  def test_initialize
    assert(@row.cells.empty?, "no cells by default")
    assert_equal(@row.worksheet, @ws, "has a reference to the worksheet")
    assert_nil(@row.height, "height defaults to nil")
    assert(!@row.custom_height?, "no custom height by default")
  end

  def test_initialize_with_fixed_height
    row = @ws.add_row([1,2,3,4,5], :height=>40)
    assert_equal(40, row.height)
    assert(row.custom_height?)
  end

  def test_style
    r = @ws.add_row([1,2,3,4,5])
    r.style=1
    r.cells.each { |c| assert_equal(c.style,1) }
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

  def test_custom_height
    @row.height = 20
    assert(@row.custom_height?)
  end

  def test_height
    assert_raise(ArgumentError) { @row.height = -3 }
    assert_nothing_raised { @row.height = 15 }
    assert_equal(15, @row.height)
  end

  def test_to_xml_without_custom_height
    xml = Nokogiri::XML::Builder.new
    @row.to_xml(xml)
    doc = Nokogiri::XML.parse(xml.to_xml)
    assert_equal(0, doc.xpath(".//row[@ht]").size)
    assert_equal(0, doc.xpath(".//row[@customHeight]").size)
  end

  def test_to_xml_with_custom_height
    @row.height = 20
    xml = Nokogiri::XML::Builder.new
    @row.to_xml(xml)
    doc = Nokogiri::XML.parse(xml.to_xml)
    assert_equal(1, doc.xpath(".//row[@ht=20][@customHeight=1]").size)
  end

end
