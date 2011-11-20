require 'test/unit'
require 'axlsx.rb'

class TestOverride < Test::Unit::TestCase
  def setup    
  end
  def teardown
  end
  def test_initialization_requires_Extension_and_ContentType
    err = "requires PartName and ContentType options"
    assert_raise(ArgumentError, err) { Axlsx::Override.new }
    assert_raise(ArgumentError, err) { Axlsx::Override.new :PartName=>"xml" }
    assert_raise(ArgumentError, err) { Axlsx::Override.new :ContentType=>"asdf" }
    assert_nothing_raised {Axlsx::Override.new :PartName=>"foo", :ContentType=>Axlsx::CHART_CT}
  end

  def test_content_type_restriction
    assert_raise(ArgumentError, "requires known content type") { Axlsx::Override.new :ContentType=>"asdf" }
  end
  
  def test_to_xml
    schema = Nokogiri::XML::Schema(File.open(Axlsx::CONTENT_TYPES_XSD))
    type = Axlsx::Override.new :PartName=>"somechart.xml", :ContentType=>Axlsx::CHART_CT
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
    assert_equal(errors.size, 0, "Override content type caused invalid content_type doc" + errors.map{ |e| e.message }.to_s)
    
  end


end
