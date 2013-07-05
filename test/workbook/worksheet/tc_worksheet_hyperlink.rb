require 'tc_helper.rb'

class TestWorksheetHyperlink < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    wb = p.workbook
    @ws = wb.add_worksheet
    @options = { :location => 'https://github.com/randym/axlsx?foo=1&bar=2', :tooltip => 'axlsx', :ref => 'A1', :display => 'AXSLX', :target => :internal }
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

  def test_target
    assert_equal(@options[:target], @a.instance_values['target'])
  end

  def test_display
    assert_equal(@options[:display], @a.display)
  end
  def test_ref
    assert_equal(@options[:ref], @a.ref)
  end

  def test_to_xml_string_with_non_external
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath("//xmlns:hyperlink[@ref='#{@a.ref}']").size, 1)
    assert_equal(doc.xpath("//xmlns:hyperlink[@tooltip='#{@a.tooltip}']").size, 1)
    assert_equal(doc.xpath("//xmlns:hyperlink[@location='#{@a.location}']").size, 1)
    assert_equal(doc.xpath("//xmlns:hyperlink[@display='#{@a.display}']").size, 1)
    assert_equal(doc.xpath("//xmlns:hyperlink[@r:id]").size, 0)
  end

  def test_to_xml_stirng_with_external
    @a.target = :external
    doc = Nokogiri::XML(@ws.to_xml_string)
    assert_equal(doc.xpath("//xmlns:hyperlink[@ref='#{@a.ref}']").size, 1)
    assert_equal(doc.xpath("//xmlns:hyperlink[@tooltip='#{@a.tooltip}']").size, 1)
    assert_equal(doc.xpath("//xmlns:hyperlink[@display='#{@a.display}']").size, 1)
    assert_equal(doc.xpath("//xmlns:hyperlink[@location='#{@a.location}']").size, 0)
    assert_equal(doc.xpath("//xmlns:hyperlink[@r:id='#{@a.relationship.Id}']").size, 1)
  end
end


