require 'tc_helper'

class TestWorkbookView < Test::Unit::TestCase

  def setup
    @options = { visibility: :hidden, minimized: true, show_horizontal_scroll: true, show_vertical_scroll: true,
                show_sheet_tabs: true, tab_ratio: 750, first_sheet: 0, active_tab: 1, x_window: 500, y_window: 400,
                window_width: 800, window_height: 600, auto_filter_date_grouping: true }
    @book_view = Axlsx::WorkbookView.new @options
  end

  def test_options_assignation
    @options.each do |key, value|
      assert_equal(value, @book_view.send(key))
    end
  end

  def test_boolean_attribute_validation
    %w(minimized show_horizontal_scroll show_vertical_scroll show_sheet_tabs auto_filter_date_grouping).each do |attr|
      assert_raise(ArgumentError, 'only booleanish allowed in boolean attributes') { @book_view.send("#{attr}=", "banana") }
      assert_nothing_raised { @book_view.send("#{attr}=", false )}
    end
  end

  def test_integer_attribute_validation
    %w(tab_ratio first_sheet active_tab x_window y_window window_width window_height).each do |attr|
      assert_raise(ArgumentError, 'only integer allowed in integer attributes') { @book_view.send("#{attr}=", "b") }
      assert_nothing_raised { @book_view.send("#{attr}=", 7 )}
    end
  end

  def test_visibility_attribute_validation
    assert_raise(ArgumentError) { @book_view.visibility = :foobar }
    assert_nothing_raised { @book_view.visibility = :hidden }
    assert_nothing_raised { @book_view.visibility = :very_hidden }
    assert_nothing_raised { @book_view.visibility = :visible }
  end

  def test_to_xml_string
    xml = @book_view.to_xml_string
    doc = Nokogiri::XML(xml)
    @options.each do |key, value|
      if value == true || value == false
        value = value ? 1 : 0
      end
      path = "workbookView[@#{Axlsx.camel(key, false)}='#{value}']"
      assert_equal(1, doc.xpath(path).size)
    end
  end
end
