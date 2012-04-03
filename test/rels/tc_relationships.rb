require 'tc_helper.rb'

class TestRelationships < Test::Unit::TestCase

  def test_valid_document
    @rels = Axlsx::Relationships.new
    schema = Nokogiri::XML::Schema(File.open(Axlsx::RELS_XSD))
    doc = Nokogiri::XML(@rels.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      puts error.message
      errors << error
    end

    @rels << Axlsx::Relationship.new(Axlsx::WORKSHEET_R, "bar")
    doc = Nokogiri::XML(@rels.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      puts error.message
      errors << error
    end

    assert(errors.size == 0)
  end

end
