require 'tc_helper.rb'

class TestStrData < Test::Unit::TestCase

  def setup
    @str_data = Axlsx::StrData.new :data => ["1", "2", "3"]
  end

  def test_to_xml_string_strLit
    str = '<?xml version="1.0" encoding="UTF-8"?>'
    str << '<c:chartSpace xmlns:c="' << Axlsx::XML_NS_C << '">'
    str << @str_data.to_xml_string
    doc = Nokogiri::XML(str)
    assert_equal(doc.xpath("//c:strLit/c:ptCount[@val=3]").size, 1)
    assert_equal(doc.xpath("//c:strLit/c:pt/c:v[text()='1']").size, 1)
  end

end
