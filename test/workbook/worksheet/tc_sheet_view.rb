# encoding: UTF-8
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../../"
require 'tc_helper.rb'

class TestSheetView < Test::Unit::TestCase
  def setup
    #inverse defaults for booleans
    @boolean_options = { :right_to_left => true, :show_formulas => true, :show_outline_symbols => true,
      :show_white_space => true, :tab_selected => true, :default_grid_color => false, :show_grid_lines => false,
      :show_row_col_headers => false, :show_ruler => false, :show_zeros => false, :window_protection => true }
    @symbol_options = { :view => :page_break_preview }
    @nil_options = { :color_id => 2, :top_left_cell => 'A2' }
    @int_0 = { :zoom_scale_normal => 100, :zoom_scale_page_layout_view => 100, :zoom_scale_sheet_layout_view => 100, :workbook_view_id => 2 }
    @int_100 = { :zoom_scale => 10 }

    @integer_options = { :color_id => 2, :workbook_view_id => 2 }.merge(@int_0).merge(@int_100)
    @string_options = { :top_left_cell => 'A2' }


    @options = @boolean_options.merge(@boolean_options).merge(@symbol_options).merge(@nil_options).merge(@int_0).merge(@int_100)

    @sv = Axlsx::SheetView.new(@options)
  end

  def test_initialize
    sv = Axlsx::SheetView.new

    @boolean_options.each do |key, value|
      assert_equal(!value, sv.send(key.to_sym), "initialized default #{key} should be #{!value}")
      assert_equal(value, @sv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end

    @nil_options.each do |key, value|
      assert_equal(nil, sv.send(key.to_sym), "initialized default #{key} should be nil")
      assert_equal(value, @sv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end

    @int_0.each do |key, value|
      assert_equal(0, sv.send(key.to_sym), "initialized default #{key} should be 0")
      assert_equal(value, @sv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end

    @int_100.each do |key, value|
      assert_equal(100, sv.send(key.to_sym), "initialized default #{key} should be 100")
      assert_equal(value, @sv.send(key.to_sym), "initialized options #{key} should be #{value}")
    end
  end

  def test_boolean_attribute_validation
    @boolean_options.each do |key, value|
      assert_raise(ArgumentError, "#{key} must be boolean") { @sv.send("#{key}=".to_sym, 'A') }
      assert_nothing_raised { @sv.send("#{key}=".to_sym, true) }
    end
  end

  def test_string_attribute_validation
    @string_options.each do |key, value|
      assert_raise(ArgumentError, "#{key} must be string") { @sv.send("#{key}=".to_sym, :symbol) }
      assert_nothing_raised { @sv.send("#{key}=".to_sym, "foo") }
    end
  end

  def test_symbol_attribute_validation
    @symbol_options.each do |key, value|
      assert_raise(ArgumentError, "#{key} must be symbol") { @sv.send("#{key}=".to_sym, "foo") }
      assert_nothing_raised { @sv.send("#{key}=".to_sym, value) }
    end
  end

  def test_integer_attribute_validation
    @integer_options.each do |key, value|
      assert_raise(ArgumentError, "#{key} must be integer") { @sv.send("#{key}=".to_sym, "foo") }
      assert_nothing_raised { @sv.send("#{key}=".to_sym, value) }
    end
  end

  def test_color_id
    assert_raise(ArgumentError) { @sv.color_id = "10" }
    assert_nothing_raised { @sv.color_id = 2 }
    assert_equal(@sv.color_id, 2)
  end

  def test_default_grid_color
    assert_raise(ArgumentError) { @sv.default_grid_color = "foo" }
    assert_nothing_raised { @sv.default_grid_color = false }
    assert_equal(@sv.default_grid_color, false)
  end

  def test_right_to_left
    assert_raise(ArgumentError) { @sv.right_to_left = "foo´" }
    assert_nothing_raised { @sv.right_to_left = true }
    assert_equal(@sv.right_to_left, true)
  end

  def test_show_formulas
    assert_raise(ArgumentError) { @sv.show_formulas = 'foo' }
    assert_nothing_raised { @sv.show_formulas = false }
    assert_equal(@sv.show_formulas, false)
  end

  def test_show_grid_lines
    assert_raise(ArgumentError) { @sv.show_grid_lines = "foo" }
    assert_nothing_raised { @sv.show_grid_lines = false }
    assert_equal(@sv.show_grid_lines, false)
  end

  def test_show_outline_symbols
    assert_raise(ArgumentError) { @sv.show_outline_symbols = 'foo' }
    assert_nothing_raised { @sv.show_outline_symbols = false }
    assert_equal(@sv.show_outline_symbols, false)
  end

  def test_show_row_col_headers
    assert_raise(ArgumentError) { @sv.show_row_col_headers = "foo" }
    assert_nothing_raised { @sv.show_row_col_headers = false }
    assert_equal(@sv.show_row_col_headers, false)
  end

  def test_show_ruler
    assert_raise(ArgumentError) { @sv.show_ruler = 'foo' }
    assert_nothing_raised { @sv.show_ruler = false }
    assert_equal(@sv.show_ruler, false)
  end

  def test_show_white_space
    assert_raise(ArgumentError) { @sv.show_white_space = 'foo' }
    assert_nothing_raised { @sv.show_white_space = false }
    assert_equal(@sv.show_white_space, false)
  end

  def test_show_zeros
    assert_raise(ArgumentError) { @sv.show_zeros = "foo" }
    assert_nothing_raised { @sv.show_zeros = false }
    assert_equal(@sv.show_zeros, false)
  end

  def test_tab_selected
    assert_raise(ArgumentError) { @sv.tab_selected = "foo" }
    assert_nothing_raised { @sv.tab_selected = false }
    assert_equal(@sv.tab_selected, false)
  end

  def test_top_left_cell
    assert_raise(ArgumentError) { @sv.top_left_cell = :cell_adress }
    assert_nothing_raised { @sv.top_left_cell = "A2" }
    assert_equal(@sv.top_left_cell, "A2")
  end

  def test_view
    assert_raise(ArgumentError) { @sv.view = 'view' }
    assert_nothing_raised { @sv.view = :page_break_preview }
    assert_equal(@sv.view, :page_break_preview)
  end

  def test_window_protection
    assert_raise(ArgumentError) { @sv.window_protection = "foo" }
    assert_nothing_raised { @sv.window_protection = false }
    assert_equal(@sv.window_protection, false)
  end

  def test_workbook_view_id
    assert_raise(ArgumentError) { @sv.workbook_view_id = "1" }
    assert_nothing_raised { @sv.workbook_view_id = 1 }
    assert_equal(@sv.workbook_view_id, 1)
  end

  def test_zoom_scale
    assert_raise(ArgumentError) { @sv.zoom_scale = "50" }
    assert_nothing_raised { @sv.zoom_scale = 50 }
    assert_equal(@sv.zoom_scale, 50)
  end

  def test_zoom_scale_normal
    assert_raise(ArgumentError) { @sv.zoom_scale_normal = "50" }
    assert_nothing_raised { @sv.zoom_scale_normal = 50 }
    assert_equal(@sv.zoom_scale_normal, 50)
  end

  def test_zoom_scale_page_layout_view
    assert_raise(ArgumentError) { @sv.zoom_scale_page_layout_view = "50" }
    assert_nothing_raised { @sv.zoom_scale_page_layout_view = 50 }
    assert_equal(@sv.zoom_scale_page_layout_view, 50)
  end

  def test_zoom_scale_sheet_layout_view
    assert_raise(ArgumentError) { @sv.zoom_scale_sheet_layout_view = "50" }
    assert_nothing_raised { @sv.zoom_scale_sheet_layout_view = 50 }
    assert_equal(@sv.zoom_scale_sheet_layout_view, 50)
  end

  def test_to_xml
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name => "sheetview"
    @ws.sheet_view do |vs|
      vs.view = :page_break_preview
    end

    doc = Nokogiri::XML.parse(@ws.sheet_view.to_xml_string)

    assert_equal(1, doc.xpath("//sheetView[@tabSelected='false']").size)

    assert_equal(1, doc.xpath("//sheetView[@tabSelected='false'][@showWhiteSpace='false'][@showOutlineSymbols='false'][@showFormulas='false']
        [@rightToLeft='false'][@windowProtection='false'][@showZeros='true'][@showRuler='true']
        [@showRowColHeaders='true'][@showGridLines='true'][@defaultGridColor='true']
        [@zoomScale='100'][@workbookViewId='0'][@zoomScaleSheetLayoutView='0'][@zoomScalePageLayoutView='0']
        [@zoomScaleNormal='0'][@view='pageBreakPreview']").size)
  end

  def test_add_selection
     @sv.add_selection(:top_left, :active_cell => "A1")
     assert_equal('A1', @sv.selections[:top_left].active_cell)
  end

end
