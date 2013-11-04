 require 'tc_helper.rb'

 class TestNumDataSource < Test::Unit::TestCase

  def setup
    @data_source = Axlsx::NumDataSource.new :data => ["1", "2", "3"]
  end

  def test_tag_name
    assert_raise(ArgumentError) { @data_source.tag_name = :zVal }
    assert_nothing_raised { @data_source.tag_name = :yVal }
    assert_nothing_raised { @data_source.tag_name = :bubbleSize }
  end

  def test_to_xml_string_strLit
    str = '<?xml version="1.0" encoding="UTF-8"?>'
    str << '<c:chartSpace xmlns:c="' << Axlsx::XML_NS_C << '">'
    str << @data_source.to_xml_string
    doc = Nokogiri::XML(str)
    assert_equal(doc.xpath("//c:val").size, 1)
  end

 end
