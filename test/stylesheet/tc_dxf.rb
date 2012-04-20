require 'tc_helper.rb'

class TestDxf < Test::Unit::TestCase

  def setup
    @item = Axlsx::Dxf.new
    @styles = Axlsx::Styles.new    
  end

  def teardown
  end

  def test_initialiation
    assert_equal(@item.alignment, nil)
    assert_equal(@item.protection, nil)
    assert_equal(@item.numFmt, nil)
    assert_equal(@item.font, nil)
    assert_equal(@item.fill, nil)
    assert_equal(@item.border, nil)
  end

  def test_alignment
    assert_raise(ArgumentError) { @item.alignment = -1.1 }
    assert_nothing_raised { @item.alignment = Axlsx::CellAlignment.new }
    assert(@item.alignment.is_a?(Axlsx::CellAlignment))
  end

  def test_protection
    assert_raise(ArgumentError) { @item.protection = -1.1 }
    assert_nothing_raised { @item.protection = Axlsx::CellProtection.new }
    assert(@item.protection.is_a?(Axlsx::CellProtection))
  end

  def test_numFmt
    assert_raise(ArgumentError) { @item.numFmt = 1 }
    assert_nothing_raised { @item.numFmt = Axlsx::NumFmt.new }
    assert @item.numFmt.is_a? Axlsx::NumFmt
  end

  def test_fill
    assert_raise(ArgumentError) { @item.fill =  1 }
    assert_nothing_raised { @item.fill = Axlsx::Fill.new(Axlsx::PatternFill.new(:patternType =>:solid, :fgColor=> Axlsx::Color.new(:rgb => "FF000000"))) }
    assert @item.fill.is_a? Axlsx::Fill
  end

  def test_font
    assert_raise(ArgumentError) { @item.font = 1 }
    assert_nothing_raised { @item.font = Axlsx::Font.new }
    assert @item.font.is_a? Axlsx::Font 
  end

  def test_border
    assert_raise(ArgumentError) { @item.border = 1 }
    assert_nothing_raised { @item.border = Axlsx::Border.new }
    assert @item.border.is_a? Axlsx::Border
  end

  def test_to_xml
    @item.border = Axlsx::Border.new
    doc = Nokogiri::XML.parse(@item.to_xml_string)
    assert_equal(1, doc.xpath(".//dxf//border").size)
    assert_equal(0, doc.xpath(".//dxf//font").size)    
  end

  def test_many_options_xml
    @item.border = Axlsx::Border.new
    @item.alignment = Axlsx::CellAlignment.new
    @item.fill = Axlsx::Fill.new(Axlsx::PatternFill.new(:patternType =>:solid, :fgColor=> Axlsx::Color.new(:rgb => "FF000000")))
    @item.font = Axlsx::Font.new
    @item.protection = Axlsx::CellProtection.new
    @item.numFmt = Axlsx::NumFmt.new
    
    doc = Nokogiri::XML.parse(@item.to_xml_string)
    assert_equal(1, doc.xpath(".//dxf//fill//patternFill[@patternType='solid']//fgColor[@rgb='FF000000']").size)
    assert_equal(1, doc.xpath(".//dxf//font").size)
    assert_equal(1, doc.xpath(".//dxf//protection").size)
    assert_equal(1, doc.xpath(".//dxf//numFmt[@numFmtId='0'][@formatCode='']").size)
    assert_equal(1, doc.xpath(".//dxf//alignment").size)
    assert_equal(1, doc.xpath(".//dxf//border").size)    
  end
end
