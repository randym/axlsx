$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../"

require 'tc_helper.rb'

class TestTitle < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @title = Axlsx::Title.new
    @chart = ws.add_chart Axlsx::Bar3DChart
  end

  def teardown
  end

  def test_initialization
    assert(@title.text == "")
    assert(@title.cell == nil)
  end

  def test_text
    assert_raise(ArgumentError, "text must be a string") { @title.text = 123 }
    @title.cell = @row.cells.first
    @title.text = "bob"
    assert(@title.cell == nil, "setting title with text clears the cell")
  end

  def test_cell
    assert_raise(ArgumentError, "cell must be a Cell") { @title.cell = "123" }
    @title.cell = @row.cells.first
    assert(@title.text == "one")
  end

  def test_to_xml_string_text
    @title.text = 'foo'
    doc = Nokogiri::XML(@title.to_xml_string)
    assert_equal(1, doc.xpath('//rich').size)
    assert_equal(1, doc.xpath('//t[text()="foo"]').size)
  end

 def test_to_xml_string_cell
    @title.cell = @row.cells.first
    doc = Nokogiri::XML(@title.to_xml_string)
    puts doc.to_xml
    assert_equal(1, doc.xpath('//strCache').size)
    assert_equal(1, doc.xpath('//v[text()="one"]').size)
  end


end
