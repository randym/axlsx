require 'tc_helper.rb'

class TestWorksheetHyperlink < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    wb = p.workbook
    @ws = wb.add_worksheet
    @options = { :location => 'https://github.com/randym/axlsx', :tooltip => 'axlsx', :ref => 'A1', :display => 'AXSLX', :r_id => 'rId1' }
    @a = @ws.add_hyperlink @options
  end

  def test_initailize
    assert_raise(ArgumentError) { Axlsx::WorksheetHyperlink.new }
  end

  def test_location
    assert_equal(@options[:location], @a.location)
  end

  def test_tooltip
    assert_equal(@options[:tooltip], @a.tooltip)
  end

  def test_display
    assert_equal(@options[:display], @a.display)
  end
  def test_ref
    assert_equal(@options[:ref], @a.ref)
  end
  def test_r_id
    assert_equal("rId1", @a.r_id)
  end


  def test_to_xml_string
    doc = Nokogiri::XML(@a.to_xml_string)
    puts doc.to_xml
    assert_equal(doc.xpath("//hyperlink[@ref='#{@a.ref}']").size, 1)
    assert_equal(doc.xpath("//hyperlink[@tooltip='#{@a.tooltip}']").size, 1)
    assert_equal(doc.xpath("//hyperlink[@display='#{@a.display}']").size, 1)
    assert_equal(doc.xpath("//hyperlink[@location='#{@a.location}']").size, 1)
  end
end


