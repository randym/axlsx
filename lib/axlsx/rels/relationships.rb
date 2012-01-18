module Axlsx
require 'axlsx/rels/relationship.rb'

  # Relationships are a collection of Relations that define how package parts are related.
  # @note The package automatically manages releationships.
  class Relationships < SimpleTypedList

    # Creates a new Relationships collection based on SimpleTypedList
    def initialize
      super Relationship
    end

    # Serializes the relationships document.
    # @return [String]
    def to_xml()
      builder = Nokogiri::XML::Builder.new(:encoding => ENCODING) do |xml|
        xml.Relationships(:xmlns => Axlsx::RELS_R) {
          each_with_index { |rel, index| rel.to_xml(xml, "rId#{index+1}") }
        }
      end
      builder.to_xml(:save_with => 0)
    end
  
  end
end
