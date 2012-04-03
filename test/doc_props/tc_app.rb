require 'tc_helper.rb'

class TestApp < Test::Unit::TestCase
  def test_valid_document
    schema = Nokogiri::XML::Schema(File.open(Axlsx::APP_XSD))
    doc = Nokogiri::XML(Axlsx::App.new.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors << error
    end
    assert_equal(errors.size, 0, "app.xml invalid" + errors.map{ |e| e.message }.to_s)
  end
end
