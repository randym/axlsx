require 'tc_helper.rb'

class TestApp < Test::Unit::TestCase
  def setup
    options = {
      :'Template' => 'Foo.xlt',
      :'Manager' => 'Penny',
      :'Company' => "Bob's Repair",
      :'Pages' => 1,
      :'Words' => 2,
      :'Characters' => 7,
      :'PresentationFormat' => 'any',
      :'Lines' => 1,
      :'Paragraphs' => 1,
      :'Slides' => 4,
      :'Notes' => 1,
      :'TotalTime' => 2,
      :'HidddenSlides' => 3,
      :'MMClips' => 10,
      :'ScaleCrop' => true,
      :'LinksUpToDate' => true,
      :'CharactersWithSpaces' => 9,
      :'SharedDoc' => false,
      :'HyperlinkBase' => 'foo',
      :'HyperlInksChanged' => false, 
      :'Application' => 'axlsx',
      :'AppVersion' => '1.1.5',
      :'DocSecurity' => 0
      }

    @app = Axlsx::App.new options

  end
  def test_valid_document
    schema = Nokogiri::XML::Schema(File.open(Axlsx::APP_XSD))
    doc = Nokogiri::XML(@app.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors << error
    end
    assert_equal(errors.size, 0, "app.xml invalid" + errors.map{ |e| e.message }.to_s)
  end
end
