require 'tc_helper.rb'

class TestSheetFormatPr < Test::Unit::TestCase

  def setup
    @options = {
      :base_col_width => 5,
      :default_col_width => 7.2,
      :default_row_height => 5.2,
      :custom_height => true,
      :zero_height => false,
      :thick_top => true,
      :thick_bottom => true,
      :outline_level_row => 0,
      :outline_level_col => 0
    }
    @sheet_format_pr = Axlsx::SheetFormatPr.new(@options)
  end

  def test_default_initialization
    sheet_format_pr = Axlsx::SheetFormatPr.new
    assert_equal 8, sheet_format_pr.base_col_width
    assert_equal 18, sheet_format_pr.default_row_height
  end

  def test_initialization_with_options
    @options.each do |key, value|
      assert_equal value, @sheet_format_pr.instance_variable_get("@#{key}")
    end
  end

  def test_base_col_width
    assert_raise(ArgumentError) { @sheet_format_pr.base_col_width = :foo }
    assert_nothing_raised { @sheet_format_pr.base_col_width = 1 }
  end

  def test_outline_level_row
    assert_raise(ArgumentError) { @sheet_format_pr.outline_level_row = :foo }
    assert_nothing_raised { @sheet_format_pr.outline_level_row = 1 }
  end

  def test_outline_level_col
    assert_raise(ArgumentError) { @sheet_format_pr.outline_level_col = :foo }
    assert_nothing_raised { @sheet_format_pr.outline_level_col = 1 }
  end

  def test_default_row_height
   assert_raise(ArgumentError) { @sheet_format_pr.default_row_height = :foo }
   assert_nothing_raised { @sheet_format_pr.default_row_height= 1.0 }
  end

  def test_default_col_width
   assert_raise(ArgumentError) { @sheet_format_pr.default_col_width= :foo }
   assert_nothing_raised { @sheet_format_pr.default_col_width = 1.0 }
  end

  def test_custom_height
   assert_raise(ArgumentError) { @sheet_format_pr.custom_height= :foo }
   assert_nothing_raised { @sheet_format_pr.custom_height = true }
  end

  def test_zero_height
   assert_raise(ArgumentError) { @sheet_format_pr.zero_height= :foo }
   assert_nothing_raised { @sheet_format_pr.zero_height = true }
  end
  def test_thick_top
   assert_raise(ArgumentError) { @sheet_format_pr.thick_top= :foo }
   assert_nothing_raised { @sheet_format_pr.thick_top = true }
  end
  def test_thick_bottom
   assert_raise(ArgumentError) { @sheet_format_pr.thick_bottom= :foo }
   assert_nothing_raised { @sheet_format_pr.thick_bottom = true }
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@sheet_format_pr.to_xml_string)
    assert doc.xpath("sheetFormatPr[@thickBottom=1]")
    assert doc.xpath("sheetFormatPr[@baseColWidth=5]")
    assert doc.xpath("sheetFormatPr[@default_col_width=7.2]")
    assert doc.xpath("sheetFormatPr[@default_row_height=5.2]")
    assert doc.xpath("sheetFormatPr[@custom_height=1]")
    assert doc.xpath("sheetFormatPr[@zero_height=0]")
    assert doc.xpath("sheetFormatPr[@thick_top=1]")
    assert doc.xpath("sheetFormatPr[@outline_level_row=0]")
    assert doc.xpath("sheetFormatPr[@outline_level_col=0]")
  end

end
