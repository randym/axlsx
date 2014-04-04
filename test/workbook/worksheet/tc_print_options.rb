require 'tc_helper.rb'

class TestPrintOptions < Test::Unit::TestCase

  def setup
    p = Axlsx::Package.new
    ws = p.workbook.add_worksheet :name => "hmmm"
    @po = ws.print_options
  end

  def test_initialize
    assert_equal(false, @po.grid_lines)
    assert_equal(false, @po.headings)
    assert_equal(false, @po.horizontal_centered)
    assert_equal(false, @po.vertical_centered)
  end

  def test_initialize_with_options
    optioned = Axlsx::PrintOptions.new(:grid_lines => true, :headings => true, :horizontal_centered => true, :vertical_centered => true)
    assert_equal(true, optioned.grid_lines)
    assert_equal(true, optioned.headings)
    assert_equal(true, optioned.horizontal_centered)
    assert_equal(true, optioned.vertical_centered)
  end

  def test_set_all_values
    @po.set(:grid_lines => true, :headings => true, :horizontal_centered => true, :vertical_centered => true)
    assert_equal(true, @po.grid_lines)
    assert_equal(true, @po.headings)
    assert_equal(true, @po.horizontal_centered)
    assert_equal(true, @po.vertical_centered)
  end

  def test_set_some_values
    @po.set(:grid_lines => true, :headings => true)
    assert_equal(true, @po.grid_lines)
    assert_equal(true, @po.headings)
    assert_equal(false, @po.horizontal_centered)
    assert_equal(false, @po.vertical_centered)
  end

  def test_to_xml
    @po.set(:grid_lines => true, :headings => true, :horizontal_centered => true, :vertical_centered => true)
    doc = Nokogiri::XML.parse(@po.to_xml_string)
    assert_equal(1, doc.xpath(".//printOptions[@gridLines=1][@headings=1][@horizontalCentered=1][@verticalCentered=1]").size)
  end

  def test_grid_lines
    assert_raise(ArgumentError) { @po.grid_lines = 99 }
    assert_nothing_raised { @po.grid_lines = true }
    assert_equal(@po.grid_lines, true)
  end

  def test_headings
    assert_raise(ArgumentError) { @po.headings = 99 }
    assert_nothing_raised { @po.headings = true }
    assert_equal(@po.headings, true)
  end

  def test_horizontal_centered
    assert_raise(ArgumentError) { @po.horizontal_centered = 99 }
    assert_nothing_raised { @po.horizontal_centered = true }
    assert_equal(@po.horizontal_centered, true)
  end

  def test_vertical_centered
    assert_raise(ArgumentError) { @po.vertical_centered = 99 }
    assert_nothing_raised { @po.vertical_centered = true }
    assert_equal(@po.vertical_centered, true)
  end

end
