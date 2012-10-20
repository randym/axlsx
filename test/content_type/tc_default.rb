# encoding: UTF-8
require 'tc_helper.rb'

class TestDefault < Test::Unit::TestCase

  def test_content_type_restriction
    assert_raise(ArgumentError, "raises argument error if invlalid ContentType is") { Axlsx::Default.new :ContentType=>"asdf" }
  end

  def test_to_xml_string
    type = Axlsx::Default.new :Extension=>"xml", :ContentType=>Axlsx::XML_CT
    doc = Nokogiri::XML(type.to_xml_string)
    assert_equal(doc.xpath("Default[@ContentType='#{Axlsx::XML_CT}']").size, 1)
    assert_equal(doc.xpath("Default[@Extension='xml']").size, 1)
  end
end
