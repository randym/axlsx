require 'test/unit'
require 'axlsx.rb'

class TestApp < Test::Unit::TestCase
  def setup    
  end
  def teardown
  end
  
  def test_valid_document
    schema = Nokogiri::XML::Schema(File.open(Axlsx::APP_XSD))
    doc = Nokogiri::XML(Axlsx::App.new.to_xml)
    errors = []
    schema.validate(doc).each do |error|
      errors << error
    end
    assert_equal(errors.size, 0, "app.xml invalid" + errors.map{ |e| e.message }.to_s)
  end
end
