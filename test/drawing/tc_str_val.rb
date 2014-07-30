require 'tc_helper.rb'

class TestStrVal < Minitest::Unit::TestCase

  def setup
    @str_val = Axlsx::StrVal.new :v => "1"
  end

  def test_initialize
    assert_equal(@str_val.v, "1")
  end

  def test_to_xml_string
    str = '<?xml version="1.0" encoding="UTF-8"?>'
    str << '<c:chartSpace xmlns:c="' << Axlsx::XML_NS_C << '">'
    str << @str_val.to_xml_string(0)
    doc = Nokogiri::XML(str)
    assert_equal(doc.xpath("//c:pt/c:v[text()='1']").size, 1)
  end

end
