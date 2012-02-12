# encoding: UTF-8

require 'test/unit'
require 'axlsx.rb'

class TestDefault < Test::Unit::TestCase
  def setup    
  end
  def teardown
  end
  def test_initialization_requires_Extension_and_ContentType
    assert_raise(ArgumentError, "raises argument error if Extension and/or ContentType are not specified") { Axlsx::Default.new }
    assert_raise(ArgumentError, "raises argument error if Extension and/or ContentType are not specified") { Axlsx::Default.new :Extension=>"xml" }
    assert_raise(ArgumentError, "raises argument error if Extension and/or ContentType are not specified") { Axlsx::Default.new :ContentType=>"asdf" }

    assert_nothing_raised {Axlsx::Default.new :Extension=>"foo", :ContentType=>Axlsx::XML_CT}

  end
  def test_content_type_restriction
    assert_raise(ArgumentError, "raises argument error if invlalid ContentType is") { Axlsx::Default.new :ContentType=>"asdf" }
  end
  
  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::CONTENT_TYPES_XSD))
    type = Axlsx::Default.new :Extension=>"xml", :ContentType=>Axlsx::XML_CT
    builder = Nokogiri::XML::Builder.new(:encoding => Axlsx::ENCODING) do |xml|
      xml.Types(:xmlns => Axlsx::XML_NS_T) {
        type.to_xml(xml)
      }
    end
    doc = Nokogiri::XML(builder.to_xml)
    errors = []
    schema.validate(doc).each do |error|
      puts error.message
      errors << error
    end
    assert_equal(errors.size, 0, "[Content Types].xml Invalid" + errors.map{ |e| e.message }.to_s)
    
  end


end
